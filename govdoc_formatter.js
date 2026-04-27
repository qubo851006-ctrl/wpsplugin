(function initGovDocFormatter() {
  "use strict";

  const DEFAULT_RULES = {
    standard: "GB/T 9704-2012",
    page: { paperSize: "A4", topCm: 3.7, bottomCm: 3.5, leftCm: 2.8, rightCm: 2.6 },
    body: { fontName: "仿宋_GB2312", fontSizePt: 16, lineSpacingMultiple: 1.4, firstLineIndentChars: 2 },
    title: { fontName: "方正小标宋简体", fontSizePt: 22, align: "center" },
    headings: {
      level1: { pattern: "^[一二三四五六七八九十]+、", fontName: "黑体", fontSizePt: 16 },
      level2: { pattern: "^（[一二三四五六七八九十]+）", fontName: "楷体_GB2312", fontSizePt: 16 },
      level3: { pattern: "^\\d+\\.", fontName: "仿宋_GB2312", fontSizePt: 16 }
    },
    checks: {
      requireDocTypeInTitle: true,
      docTypes: ["通知", "请示", "报告", "函", "纪要", "通报", "意见", "决定", "公告", "通告"],
      structureHints: ["主送机关", "落款日期", "附件说明", "页码"]
    }
  };

  let rulesCache = null;

  function cmToPoints(cm) {
    return cm * 28.3464567;
  }

  function getApplication() {
    if (typeof window.Application !== "undefined") return window.Application;
    if (typeof Application !== "undefined") return Application;
    try {
      if (window.parent && typeof window.parent.Application !== "undefined") {
        return window.parent.Application;
      }
    } catch (e) {}
    return null;
  }

  function getActiveDocument() {
    const app = getApplication();
    if (!app || !app.ActiveDocument) return null;
    return app.ActiveDocument;
  }

  async function loadRules() {
    if (rulesCache) return rulesCache;
    try {
      const resp = await fetch("./rules_govdoc.json", { cache: "no-store" });
      if (resp.ok) {
        rulesCache = await resp.json();
        return rulesCache;
      }
    } catch (e) {}
    rulesCache = DEFAULT_RULES;
    return rulesCache;
  }

  function getParagraphItem(paragraphs, index) {
    if (!paragraphs) return null;
    try {
      if (typeof paragraphs.Item === "function") return paragraphs.Item(index);
    } catch (e) {}
    try {
      if (typeof paragraphs.item === "function") return paragraphs.item(index);
    } catch (e) {}
    try {
      if (typeof paragraphs === "function") return paragraphs(index);
    } catch (e) {}
    return null;
  }

  function getParagraphCollection(doc) {
    try {
      if (doc?.Paragraphs?.Count) return doc.Paragraphs;
    } catch (e) {}
    try {
      if (doc?.Content?.Paragraphs?.Count) return doc.Content.Paragraphs;
    } catch (e) {}
    return null;
  }

  function normalizeParagraphText(rawText) {
    return String(rawText || "")
      .replace(/[\r\x07]/g, "")
      .replace(/\u000b/g, "")
      .trim();
  }

  function collectParagraphs(doc) {
    const rows = [];
    try {
      const paragraphs = getParagraphCollection(doc);
      if (!paragraphs || !paragraphs.Count) return rows;
      for (let i = 1; i <= paragraphs.Count; i++) {
        const p = getParagraphItem(paragraphs, i);
        if (!p) continue;
        const text = normalizeParagraphText(p?.Range?.Text || "");
        rows.push({ index: i, paragraph: p, text });
      }
    } catch (e) {}
    return rows;
  }

  function firstNonEmpty(rows) {
    for (const row of rows) {
      if (row.text) return row;
    }
    return null;
  }

  function detectHeadingLevel(text, rules) {
    if (!text) return "";
    const h = rules.headings || {};
    try {
      if (h.level1?.pattern && new RegExp(h.level1.pattern).test(text)) return "level1";
      if (h.level2?.pattern && new RegExp(h.level2.pattern).test(text)) return "level2";
      if (h.level3?.pattern && new RegExp(h.level3.pattern).test(text)) return "level3";
    } catch (e) {}
    return "";
  }

  function getParagraphFontName(paragraph) {
    try {
      return paragraph.Range.Font.NameFarEast || paragraph.Range.Font.Name || "";
    } catch (e) {
      return "";
    }
  }

  function getParagraphFontSize(paragraph) {
    try {
      return Number(paragraph.Range.Font.Size || 0);
    } catch (e) {
      return 0;
    }
  }

  function setParagraphStyle(paragraph, spec, options = {}) {
    if (!paragraph || !spec) return false;
    let changed = false;
    try {
      if (spec.fontName) {
        paragraph.Range.Font.NameFarEast = spec.fontName;
        paragraph.Range.Font.Name = spec.fontName;
        changed = true;
      }
    } catch (e) {}

    try {
      if (spec.fontSizePt) {
        paragraph.Range.Font.Size = spec.fontSizePt;
        changed = true;
      }
    } catch (e) {}

    try {
      if (typeof options.align === "number") {
        paragraph.Range.ParagraphFormat.Alignment = options.align;
        changed = true;
      }
    } catch (e) {}

    try {
      paragraph.Range.ParagraphFormat.SpaceBefore = 0;
      paragraph.Range.ParagraphFormat.SpaceAfter = 0;
      changed = true;
    } catch (e) {}

    return changed;
  }

  function setBodyParagraphLayout(paragraph, rules) {
    let changed = false;
    try {
      const pf = paragraph.Range.ParagraphFormat;
      const firstLineIndentPt = (rules.body.firstLineIndentChars || 2) * (rules.body.fontSizePt || 16);
      pf.FirstLineIndent = firstLineIndentPt;
      changed = true;
    } catch (e) {}

    const spacingMultiple = Number(rules.body.lineSpacingMultiple || 1.4);
    try {
      const pf = paragraph.Range.ParagraphFormat;
      if (setLineSpacingMultiple(pf, spacingMultiple)) changed = true;
    } catch (e) {}
    return changed;
  }

  function setLineSpacingMultiple(paragraphFormat, multiple) {
    if (!paragraphFormat) return false;
    const app = getApplication();
    const points = app && typeof app.LinesToPoints === "function"
      ? app.LinesToPoints(multiple)
      : multiple * 12;
    let ok = false;

    // Strategy A: Word-compatible multiple spacing in points.
    try {
      paragraphFormat.LineSpacingRule = 5; // wdLineSpaceMultiple
      paragraphFormat.LineSpacing = points;
      ok = true;
    } catch (e) {}

    // Strategy B: Some WPS builds treat LineSpacing directly as multiplier.
    if (!ok) {
      try {
        paragraphFormat.LineSpacingRule = 5;
      } catch (e) {}
      try {
        paragraphFormat.LineSpacing = multiple;
        ok = true;
      } catch (e) {}
    }

    // Strategy C: Fallback without rule assignment.
    if (!ok) {
      try {
        paragraphFormat.LineSpacing = points;
        ok = true;
      } catch (e) {}
    }

    // Normalize paragraph spacing to avoid visual mismatch.
    try {
      paragraphFormat.SpaceBefore = 0;
      paragraphFormat.SpaceAfter = 0;
    } catch (e) {}
    try {
      paragraphFormat.LineUnitBefore = 0;
      paragraphFormat.LineUnitAfter = 0;
    } catch (e) {}

    return ok;
  }

  function setRangeBodyLayout(range, rules) {
    let changed = false;
    if (!range) return changed;
    try {
      const pf = range.ParagraphFormat;
      if (setLineSpacingMultiple(pf, Number(rules.body.lineSpacingMultiple || 1.4))) changed = true;
    } catch (e) {}
    try {
      const firstLineIndentPt = (rules.body.firstLineIndentChars || 2) * (rules.body.fontSizePt || 16);
      range.ParagraphFormat.FirstLineIndent = firstLineIndentPt;
      changed = true;
    } catch (e) {}
    return changed;
  }

  function getSection(doc, i) {
    let s = null;
    try { s = doc.Sections(i); } catch(e) {}
    if (!s) { try { s = doc.Sections.Item(i); } catch(e) {} }
    if (!s) { try { s = doc.Sections.Item(i - 1); } catch(e) {} }
    return s;
  }

  function getFooter(section, type) {
    let f = null;
    try { f = section.Footers(type); } catch(e) {}
    if (!f) { try { f = section.Footers.Item(type); } catch(e) {} }
    if (!f) { try { f = section.Footers.Item(type - 1); } catch(e) {} }
    return f;
  }

  function hasPageNumber(doc) {
    try {
      const secCount = doc?.Sections?.Count || 0;
      for (let i = 1; i <= secCount; i++) {
        const section = getSection(doc, i);
        if (!section) continue;
        for (const footerType of [1, 2, 3]) {
          try {
            const footer = getFooter(section, footerType);
            if (!footer) continue;
            if (footer?.PageNumbers?.Count > 0) return true;
            if (footer?.Range?.Fields?.Count > 0) return true;
            const txt = String(footer?.Range?.Text || "").trim();
            if (txt && (txt.includes("PAGE") || /\d+/.test(txt))) return true;
          } catch (e) {}
        }
      }
    } catch (e) {}
    return false;
  }

  function resolveAlign(alignName, fallback) {
    if (alignName === "left") return 0;
    if (alignName === "center") return 1;
    if (alignName === "right") return 2;
    return fallback;
  }

  function setBooleanCompat(target, prop, value) {
    if (!target) return false;
    const tries = value ? [true, -1, 1] : [false, 0];
    for (const v of tries) {
      try {
        target[prop] = v;
        const read = target[prop];
        if (value) {
          if (read === true || read === -1 || read === 1) return true;
        } else if (read === false || read === 0) {
          return true;
        }
      } catch (e) {}
    }
    return false;
  }

  function setFooterAlignment(range, alignName) {
    if (!range?.ParagraphFormat) return false;
    const candidates = alignName === "left"
      ? [0, 1]
      : alignName === "center"
      ? [1, 2]
      : [2, 3];
    for (const v of candidates) {
      try {
        range.ParagraphFormat.Alignment = v;
        const read = Number(range.ParagraphFormat.Alignment);
        if (read === v) return true;
      } catch (e) {}
    }
    return false;
  }

  function applyFooterFont(range, rules) {
    try {
      range.Font.NameFarEast = rules?.pageNumber?.fontName || "宋体";
      range.Font.Name = rules?.pageNumber?.fontName || "宋体";
      range.Font.Size = Number(rules?.pageNumber?.fontSizePt || 14);
    } catch (e) {}
  }

  function clearRangeText(range) {
    try {
      range.Text = "";
      return true;
    } catch (e) {}
    try {
      const txt = String(range.Text || "");
      if (!txt) return true;
      range.Delete();
      return true;
    } catch (e2) {}
    return false;
  }

  function appendTextToRange(range, text) {
    try {
      range.InsertAfter(text);
      return true;
    } catch (e) {}
    try {
      range.Text = String(range.Text || "") + text;
      return true;
    } catch (e2) {}
    return false;
  }

  function appendPageFieldToRange(range) {
    try {
      if (range?.Fields?.Add) {
        range.Fields.Add(range, 33); // wdFieldPage
        return true;
      }
    } catch (e) {}
    return false;
  }

  function clearFooterPageArtifacts(footer) {
    if (!footer?.Range) return false;
    let changed = false;
    const range = footer.Range;

    try {
      const pageNumbers = footer.PageNumbers;
      if (pageNumbers?.Count) {
        for (let i = pageNumbers.Count; i >= 1; i--) {
          try {
            const item = pageNumbers.Item ? pageNumbers.Item(i) : pageNumbers(i);
            if (item?.Delete) item.Delete();
            changed = true;
          } catch (e) {}
        }
      }
    } catch (e) {}

    try {
      const fields = range.Fields;
      if (fields?.Count) {
        for (let i = fields.Count; i >= 1; i--) {
          try {
            const f = fields.Item ? fields.Item(i) : fields(i);
            if (f?.Delete) {
              f.Delete();
              changed = true;
            }
          } catch (e) {}
        }
      }
    } catch (e) {}

    try {
      const txt = String(range.Text || "").replace(/[\r\x07]/g, "").trim();
      if (txt && /^[-—\s\d]+$/.test(txt)) {
        range.Text = "";
        changed = true;
      }
    } catch (e) {}

    return changed;
  }

  function writeStandardPageNumber(footer, align, rules) {
    if (!footer?.Range) return false;
    const range = footer.Range;

    const alignName = align === 0 ? "left" : align === 1 ? "center" : "right";

    const withDash = rules?.pageNumber?.withDash !== false;
    const leftDash = withDash ? "—" : "";
    const rightDash = withDash ? "—" : "";

    let ok = false;
    try {
      clearRangeText(range);
      // 清空文本后再设置对齐，防止 clearRangeText 重置段落格式
      setFooterAlignment(range, alignName);
      // 清除段落缩进，防止页码被推离左/右边缘
      try {
        range.ParagraphFormat.LeftIndent = 0;
        range.ParagraphFormat.FirstLineIndent = 0;
        range.ParagraphFormat.RightIndent = 0;
      } catch (e) {}
      applyFooterFont(range, rules);
      appendTextToRange(range, leftDash);
      ok = appendPageFieldToRange(range);
      // 域插入后 range 引用位置可能偏移，用 footer.Range.InsertAfter 保证追加到页脚末尾
      try { footer.Range.InsertAfter(rightDash); } catch(e) { appendTextToRange(range, rightDash); }
    } catch (e) {}

    if (!ok) {
      try {
        if (footer?.PageNumbers?.Add) {
          footer.PageNumbers.Add();
          ok = true;
        }
      } catch (e) {}
      if (ok) {
        // 备用路径：写入后重新设置对齐和破折号
        // 注意：不能用 range.Text = "—N—" 会覆盖掉动态域，改用 InsertBefore/InsertAfter
        try { setFooterAlignment(range, alignName); } catch (e) {}
        if (withDash) {
          try { range.InsertBefore(leftDash); } catch (e) {}
          try { footer.Range.InsertAfter(rightDash); } catch (e) {}
        }
      }
    }

    return ok;
  }

  function addOutsidePageNumber(section, rules) {
    if (!section?.Footers) return false;
    let added = false;
    try {
      const footer = section.Footers(1) || section.Footers.Item?.(1);
      if (!footer) return false;
      clearFooterPageArtifacts(footer);

      // Try "outside" alignment constants across Office/WPS variants.
      const alignCandidates = [4, 5, 3];
      for (const align of alignCandidates) {
        try {
          if (footer?.PageNumbers?.Add) {
            footer.PageNumbers.Add(align, true);
            if ((footer.PageNumbers.Count || 0) > 0) {
              added = true;
              break;
            }
          }
        } catch (e) {}
      }

      if (!added) return false;

      try {
        footer.Range.ParagraphFormat.SpaceBefore = 0;
        footer.Range.ParagraphFormat.SpaceAfter = 0;
      } catch (e) {}
      applyFooterFont(footer.Range, rules);
      return true;
    } catch (e) {
      return false;
    }
  }

  function addPageNumberBySelection(doc, rules) {
    const app = getApplication();
    if (!app || !doc) return false;
    try {
      const win = doc.ActiveWindow || app.ActiveWindow;
      const view = win?.ActivePane?.View || win?.View;
      if (!view) return false;
      const oldSeek = view.SeekView;

      try {
        // 10: wdSeekCurrentPageFooter in Word object model
        view.SeekView = 10;
      } catch (e) {
        return false;
      }

      let inserted = false;
        try {
          const sel = app.Selection;
          if (!sel) return false;
          try { setFooterAlignment(sel.Range, rules?.pageNumber?.oddAlign || "right"); } catch (e) {}
          try {
            sel.Font.NameFarEast = rules?.pageNumber?.fontName || "宋体";
            sel.Font.Name = rules?.pageNumber?.fontName || "宋体";
            sel.Font.Size = Number(rules?.pageNumber?.fontSizePt || 14);
          } catch (e) {}

          const withDash = rules?.pageNumber?.withDash !== false;
          if (withDash && sel?.TypeText) {
            try { sel.TypeText("—"); } catch (e) {}
          }

          if (sel?.Fields?.Add) {
            try {
              // 33: wdFieldPage
              sel.Fields.Add(sel.Range, 33);
              inserted = true;
            } catch (e) {}
          }
          if (!inserted && sel?.InsertAfter) {
            sel.InsertAfter("1");
            inserted = true;
          }
          if (withDash && inserted) {
            try {
              if (sel?.TypeText) sel.TypeText("—");
              else if (sel?.InsertAfter) sel.InsertAfter(" —");
            } catch (e) {}
          }
        } finally {
          try { view.SeekView = oldSeek; } catch (e) {}
        }
      return inserted;
    } catch (e) {
      return false;
    }
  }

  function writePageNumberBySeekView(doc, rules, seekViewCandidates, alignName) {
    const app = getApplication();
    if (!app || !doc) return false;
    try {
      const win = doc.ActiveWindow || app.ActiveWindow;
      const view = win?.ActivePane?.View || win?.View;
      if (!view) return false;
      const oldSeek = view.SeekView;

      // 确保处于打印视图，否则 WPS 无法访问页脚区域
      try { view.Type = 3; } catch (e) {}

      let switched = false;
      for (const seek of seekViewCandidates) {
        try {
          view.SeekView = seek;
          // 只要赋值不抛异常即视为导航成功（WPS 读回值可能与设定值不同，不能用于验证）
          switched = true;
          break;
        } catch (e) {}
      }
      if (!switched) return false;

      try {
        const sel = app.Selection;
        if (!sel) return false;
        try {
          if (sel.Range?.Fields?.Count) {
            for (let i = sel.Range.Fields.Count; i >= 1; i--) {
              const f = sel.Range.Fields.Item ? sel.Range.Fields.Item(i) : sel.Range.Fields(i);
              if (f?.Delete) f.Delete();
            }
          }
        } catch (e) {}
        try {
          sel.WholeStory();
          sel.Delete();
        } catch (e) {}

        // 清除段落缩进，防止页码被推离左/右边缘
        try {
          sel.Range.ParagraphFormat.LeftIndent = 0;
          sel.Range.ParagraphFormat.FirstLineIndent = 0;
          sel.Range.ParagraphFormat.RightIndent = 0;
        } catch (e) {}

        try { setFooterAlignment(sel.Range, alignName); } catch (e) {}
        try {
          sel.Font.NameFarEast = rules?.pageNumber?.fontName || "宋体";
          sel.Font.Name = rules?.pageNumber?.fontName || "宋体";
          sel.Font.Size = Number(rules?.pageNumber?.fontSizePt || 14);
        } catch (e) {}

        const withDash = rules?.pageNumber?.withDash !== false;
        if (withDash && sel.TypeText) {
          try { sel.TypeText("—"); } catch (e) {}
        }
        let inserted = false;
        if (sel?.Fields?.Add) {
          try {
            sel.Fields.Add(sel.Range, 33);
            inserted = true;
          } catch (e) {}
        }
        if (!inserted && sel?.InsertAfter) {
          sel.InsertAfter("1");
          inserted = true;
        }
        if (inserted && withDash) {
          try {
            if (sel.TypeText) sel.TypeText("—");
            else if (sel.InsertAfter) sel.InsertAfter("—");
          } catch (e) {}
        }
        return inserted;
      } finally {
        try { view.SeekView = oldSeek; } catch (e) {}
      }
    } catch (e) {
      return false;
    }
  }

  function addPageNumber(doc, rules) {
    if (!rules?.pageNumber?.enabled) return 0;

    let changed = 0;
    try {
      try {
        if (doc?.PageSetup) {
          setBooleanCompat(doc.PageSetup, "OddAndEvenPagesHeaderFooter", true);
          setBooleanCompat(doc.PageSetup, "DifferentOddAndEvenPagesHeaderFooter", true);
          setBooleanCompat(doc.PageSetup, "DifferentFirstPageHeaderFooter", false);
        }
      } catch (e) {}

      const secCount = doc?.Sections?.Count || 0;
      for (let i = 1; i <= secCount; i++) {
        const section = getSection(doc, i);
        if (!section) continue;

        let oddAdded = false;
        let evenAdded = false;

        try {
          if (section?.PageSetup) {
            setBooleanCompat(section.PageSetup, "OddAndEvenPagesHeaderFooter", true);
            setBooleanCompat(section.PageSetup, "DifferentOddAndEvenPagesHeaderFooter", true);
            setBooleanCompat(section.PageSetup, "DifferentFirstPageHeaderFooter", false);
            setBooleanCompat(section.PageSetup, "MirrorMargins", true);
          }
        } catch (e) {}

        // Unlink and clear all footers first.
        for (const footerType of [1, 2, 3]) {
          try {
            const footer = getFooter(section, footerType);
            if (!footer) continue;
            try { setBooleanCompat(footer, "LinkToPrevious", false); } catch (e) {}
            if (clearFooterPageArtifacts(footer)) changed += 1;
          } catch (e) {}
        }

        // --- Odd footer (Footers(1) = primary / odd pages) ---
        try {
          const footer = getFooter(section, 1);
          if (footer) {
            const align = resolveAlign(rules?.pageNumber?.oddAlign, 2);
            if (writeStandardPageNumber(footer, align, rules)) {
              oddAdded = true;
            } else {
              const withDash = rules?.pageNumber?.withDash !== false;
              for (const a of [1, 4, 5, 3]) {
                try {
                  if (footer?.PageNumbers?.Add) {
                    footer.PageNumbers.Add(a, false);
                    if ((footer.PageNumbers.Count || 0) > 0) {
                      if (withDash) {
                        try { footer.Range.InsertBefore("—"); } catch (e) {}
                        try { footer.Range.InsertAfter("—"); } catch (e) {}
                      }
                      oddAdded = true;
                      break;
                    }
                  }
                } catch (e) {}
              }
            }
          }
        } catch (e) {}

        // --- First page footer (Footers(2) = wdHeaderFooterFirstPage) ---
        // 当 DifferentFirstPageHeaderFooter=true 时，第 1 页使用独立的 footer 2。
        // WPS 中该属性设为 false 可能不生效，所以无论如何都同步写入 footer 2，
        // 保证第 1 页必然有页码。
        try {
          const f2 = getFooter(section, 2);
          if (f2) {
            try { setBooleanCompat(f2, "LinkToPrevious", false); } catch (e) {}
            const align = resolveAlign(rules?.pageNumber?.oddAlign, 2);
            if (!writeStandardPageNumber(f2, align, rules)) {
              const withDash = rules?.pageNumber?.withDash !== false;
              for (const a of [1, 4, 5, 3]) {
                try {
                  if (f2?.PageNumbers?.Add) {
                    f2.PageNumbers.Add(a, false);
                    if ((f2.PageNumbers.Count || 0) > 0) {
                      if (withDash) {
                        try { f2.Range.InsertBefore("—"); } catch (e) {}
                        try { f2.Range.InsertAfter("—"); } catch (e) {}
                      }
                      break;
                    }
                  }
                } catch (e) {}
              }
            }
          }
        } catch (e) {}

        // --- Even footer (Footers(3) = wdHeaderFooterEvenPages) ---

        // 路径 1：writeStandardPageNumber 写入带破折号格式
        if (!evenAdded) {
          try {
            const f3 = getFooter(section, 3);
            if (f3) {
              try { setBooleanCompat(f3, "LinkToPrevious", false); } catch (e) {}
              const align = resolveAlign(rules?.pageNumber?.evenAlign, 0);
              if (writeStandardPageNumber(f3, align, rules)) evenAdded = true;
            }
          } catch (e) {}
        }

        // 路径 2：PageNumbers.Add 兜底（带破折号）
        if (!evenAdded) {
          try {
            const f3 = getFooter(section, 3);
            if (f3) {
              try { setBooleanCompat(f3, "LinkToPrevious", false); } catch (e) {}
              const withDash = rules?.pageNumber?.withDash !== false;
              for (const a of [1, 4, 5, 3]) {
                try {
                  if (f3?.PageNumbers?.Add) {
                    f3.PageNumbers.Add(a, false);
                    if ((f3.PageNumbers.Count || 0) > 0) {
                      if (withDash) {
                        try { f3.Range.InsertBefore("—"); } catch (e) {}
                        try { f3.Range.InsertAfter("—"); } catch (e) {}
                      }
                      evenAdded = true;
                      break;
                    }
                  }
                } catch (e) {}
              }
            }
          } catch (e) {}
        }

        // 路径 4：直接向 Range.Fields 写入页码域（带破折号）
        if (!evenAdded) {
          try {
            const f3 = getFooter(section, 3);
            if (f3?.Range) {
              const withDash = rules?.pageNumber?.withDash !== false;
              try { clearRangeText(f3.Range); } catch (e) {}
              try { setFooterAlignment(f3.Range, rules?.pageNumber?.evenAlign || "left"); } catch (e) {}
              try { applyFooterFont(f3.Range, rules); } catch (e) {}
              if (withDash) { try { appendTextToRange(f3.Range, "—"); } catch (e) {} }
              try {
                f3.Range.Fields.Add(f3.Range, 33);
                if ((f3.Range.Fields.Count || 0) > 0) {
                  if (withDash) { try { f3.Range.InsertAfter("—"); } catch (e) { try { appendTextToRange(f3.Range, "—"); } catch (e2) {} } }
                  evenAdded = true;
                }
              } catch (e) {}
            }
          } catch (e) {}
        }

        // 路径 5：SeekView 导航到偶数页页脚（WPS 专用兜底）
        // WPS wdSeekEvenPagesFooter = 12；Word 旧版 = 7
        if (!evenAdded) {
          try {
            const evenAlignName = rules?.pageNumber?.evenAlign || "left";
            if (writePageNumberBySeekView(doc, rules, [12, 7], evenAlignName)) {
              evenAdded = true;
            }
          } catch (e) {}
        }

        // 奇数页 SeekView 兜底（当 Footers(1) API 完全失效时）
        // WPS wdSeekOddPagesFooter = 11；Word 旧版 = 6
        if (!oddAdded) {
          try {
            const oddAlignName = rules?.pageNumber?.oddAlign || "right";
            if (writePageNumberBySeekView(doc, rules, [11, 6], oddAlignName)) {
              oddAdded = true;
            }
          } catch (e) {}
        }

        if (oddAdded) changed += 1;
        if (evenAdded) changed += 1;
      }
    } catch (e) {}

    // 全部 API 均失效时，降级为当前页脚 SeekView 写入（至少保证奇数页有页码）
    if (changed === 0) {
      try {
        if (addPageNumberBySelection(doc, rules)) changed = 1;
      } catch (e) {}
    }
    return changed;
  }

  function collectPageNumberDiagnostics(doc) {
    const out = {
      docOddEven: null,
      docDiffOddEven: null,
      sec1OddEven: null,
      sec1DiffOddEven: null,
      footer1Fields: 0,
      footer1PageNums: 0,
      footer2Fields: 0,
      footer3Fields: 0,
      footer3PageNums: 0,
      footer3Linked: null,
      footer3Text: null,
      seekView: null
    };
    try {
      out.docOddEven = doc?.PageSetup?.OddAndEvenPagesHeaderFooter ?? null;
      out.docDiffOddEven = doc?.PageSetup?.DifferentOddAndEvenPagesHeaderFooter ?? null;
    } catch (e) {}
    try {
      const app = getApplication();
      const win = doc.ActiveWindow || app?.ActiveWindow;
      const diagView = win?.ActivePane?.View || win?.View;
      out.seekView = diagView?.SeekView ?? null;
    } catch (e) {}
    try {
      const sec = doc?.Sections?.Count ? getSection(doc, 1) : null;
      if (sec?.PageSetup) {
        out.sec1OddEven = sec.PageSetup.OddAndEvenPagesHeaderFooter ?? null;
        out.sec1DiffOddEven = sec.PageSetup.DifferentOddAndEvenPagesHeaderFooter ?? null;
      }
      if (sec?.Footers) {
        for (const t of [1, 2, 3]) {
          try {
            const f = getFooter(sec, t);
            const c = f?.Range?.Fields?.Count || 0;
            out[`footer${t}Fields`] = Number(c);
            if (t === 1) out.footer1PageNums = Number(f?.PageNumbers?.Count || 0);
            if (t === 3) out.footer3PageNums = Number(f?.PageNumbers?.Count || 0);
            if (t === 3) {
              out.footer3Linked = f?.LinkToPrevious ?? null;
              const txt = String(f?.Range?.Text || "").replace(/[\r\n\x07]/g, "").trim();
              out.footer3Text = txt.slice(0, 30) || "(空)";
            }
          } catch (e) {}
        }
      }
    } catch (e) {}
    return out;
  }

  function checkDocumentTypeInTitle(titleText, rules, issues) {
    if (!rules?.checks?.requireDocTypeInTitle || !titleText) return;
    const matched = (rules.checks.docTypes || []).some((word) => titleText.includes(word));
    if (!matched) {
      issues.push({
        level: "warn",
        code: "TITLE_DOC_TYPE",
        message: "标题未识别到文种（如通知/请示/报告/函），建议补全标题文种。"
      });
    }
  }

  function checkStructureHints(allText, rules, issues) {
    for (const key of rules?.checks?.structureHints || []) {
      if (!allText.includes(key)) {
        issues.push({
          level: "info",
          code: "STRUCTURE_HINT",
          message: `未检测到“${key}”相关内容，请人工确认结构要素。`
        });
      }
    }
  }

  function checkPageNumberRule(doc, rules, issues) {
    if (!rules?.checks?.requirePageNumber) return;
    if (!hasPageNumber(doc)) {
      issues.push({
        level: "warn",
        code: "PAGE_NUMBER_MISSING",
        message: "未检测到页码，建议在页脚添加页码。"
      });
    }
  }

  function scanParagraphStyles(rows, rules, issues) {
    for (const row of rows) {
      if (!row.text) continue;
      const level = detectHeadingLevel(row.text, rules);
      if (row.index === rows[0]?.index) continue;
      const fontName = getParagraphFontName(row.paragraph);
      const fontSize = getParagraphFontSize(row.paragraph);

      if (level) {
        const spec = rules.headings[level];
        if (spec?.fontName && fontName && fontName !== spec.fontName) {
          issues.push({
            level: "warn",
            code: "HEADING_FONT",
            message: `第${row.index}段疑似${level}标题，字体为“${fontName}”，建议“${spec.fontName}”。`
          });
        }
        if (spec?.fontSizePt && fontSize && Math.abs(fontSize - spec.fontSizePt) > 0.5) {
          issues.push({
            level: "warn",
            code: "HEADING_SIZE",
            message: `第${row.index}段疑似${level}标题，字号为${fontSize}pt，建议${spec.fontSizePt}pt。`
          });
        }
      } else {
        if (rules.body.fontName && fontName && fontName !== rules.body.fontName) {
          issues.push({
            level: "warn",
            code: "BODY_FONT",
            message: `第${row.index}段正文字体为“${fontName}”，建议“${rules.body.fontName}”。`
          });
        }
      }
    }
  }

  function tryFormatPage(doc, rules) {
    let changed = 0;
    try {
      const ps = doc.PageSetup;
      ps.TopMargin = cmToPoints(rules.page.topCm);
      ps.BottomMargin = cmToPoints(rules.page.bottomCm);
      ps.LeftMargin = cmToPoints(rules.page.leftCm);
      ps.RightMargin = cmToPoints(rules.page.rightCm);
      changed += 1;
    } catch (e) {}
    return changed;
  }

  async function scan() {
    const rules = await loadRules();
    const doc = getActiveDocument();
    if (!doc) {
      return { ok: false, message: "未检测到 WPS Writer 活动文档，请在文字文档中使用。", issues: [] };
    }

    const issues = [];
    let rows = collectParagraphs(doc).filter((x) => x.text);
    // Some WPS builds expose document text but not paragraph collection cleanly.
    // Keep a fallback issue message so users can still diagnose.
    if (!rows.length) {
      try {
        const raw = normalizeParagraphText(doc?.Content?.Text || doc?.Range?.Text || "");
        if (raw) {
          return {
            ok: true,
            standard: rules.standard,
            title: "",
            paragraphCount: 0,
            issues: [{
              level: "warn",
              code: "PARAGRAPH_API_FALLBACK",
              message: "检测到文档有内容，但当前环境无法枚举段落。请从功能区按钮触发“公文体检/一键规范化”。"
            }]
          };
        }
      } catch (e) {}
    }
    const first = firstNonEmpty(rows);
    const titleText = first?.text || "";

    checkDocumentTypeInTitle(titleText, rules, issues);
    scanParagraphStyles(rows, rules, issues);
    checkStructureHints(rows.map((x) => x.text).join("\n"), rules, issues);
    checkPageNumberRule(doc, rules, issues);

    return {
      ok: true,
      standard: rules.standard,
      title: titleText,
      paragraphCount: rows.length,
      issues
    };
  }

  // 探针：测试当前 WPS 环境对页脚的写入能力，帮助诊断为什么页码加不进去
  function probeFooterCapability(doc) {
    const probe = {
      secCount: 0,
      secCallWorks: false,   // doc.Sections(1)
      secItemWorks: false,   // doc.Sections.Item(1)
      secItem0Works: false,  // doc.Sections.Item(0) - 0-based?
      f1Accessible: false,
      f1TextWritable: false,
      f1FieldAddable: false,
      f1PageNumAddable: false,
      storyRange7Accessible: false,
      storyRange6Accessible: false,
      seekView6Works: false,
      selTypeTextWorks: false,
    };
    try { probe.secCount = Number(doc?.Sections?.Count || 0); } catch(e) {}

    // 测试不同的 Section 访问语法
    let sec = null;
    try { const s = doc.Sections(1); if (s) { probe.secCallWorks = true; sec = s; } } catch(e) {}
    try { const s = doc.Sections.Item(1); if (s) { probe.secItemWorks = true; sec = sec || s; } } catch(e) {}
    try { const s = doc.Sections.Item(0); if (s) { probe.secItem0Works = true; sec = sec || s; } } catch(e) {}

    // 测试 Footers 访问
    let f1 = null;
    if (sec) {
      try { f1 = sec.Footers(1); } catch(e) {}
      if (!f1) { try { f1 = sec.Footers.Item(1); } catch(e) {} }
      if (!f1) { try { f1 = sec.Footers.Item(0); } catch(e) {} }
    }
    probe.f1Accessible = !!f1;

      if (f1) {
        // 测试 Range.Text 写入
        try {
          const before = String(f1.Range.Text || "");
          f1.Range.Text = "\u2060"; // 零宽不换行空格，写后可检测
          const after = String(f1.Range.Text || "");
          probe.f1TextWritable = after.includes("\u2060") || after !== before;
          f1.Range.Text = ""; // 清回空
        } catch (e) {}

        // 测试 Fields.Add
        try {
          const before = Number(f1.Range.Fields.Count || 0);
          f1.Range.Fields.Add(f1.Range, 33);
          const after = Number(f1.Range.Fields.Count || 0);
          probe.f1FieldAddable = after > before;
          // 清除刚加的 field
          try {
            for (let i = after; i >= before + 1; i--) {
              (f1.Range.Fields.Item?.(i) || f1.Range.Fields(i))?.Delete?.();
            }
          } catch (e) {}
        } catch (e) {}

        // 测试 PageNumbers.Add
        try {
          const before = Number(f1.PageNumbers.Count || 0);
          f1.PageNumbers.Add(1, false);
          const after = Number(f1.PageNumbers.Count || 0);
          probe.f1PageNumAddable = after > before;
          // 清除刚加的 page number
          try {
            for (let i = after; i >= before + 1; i--) {
              (f1.PageNumbers.Item?.(i) || f1.PageNumbers(i))?.Delete?.();
            }
          } catch (e) {}
        } catch (e) {}
      }

      // 测试 StoryRanges（另一种访问页脚的途径）
      try {
        const sr7 = doc.StoryRanges && doc.StoryRanges.Item(7); // primary footer
        probe.storyRange7Accessible = !!sr7;
      } catch (e) {}
      try {
        const sr6 = doc.StoryRanges && doc.StoryRanges.Item(6); // even footer
        probe.storyRange6Accessible = !!sr6;
      } catch (e) {}

      // 测试 SeekView=6 + TypeText
      try {
        const app = getApplication();
        const win = doc.ActiveWindow || app?.ActiveWindow;
        const view = win?.ActivePane?.View || win?.View;
        if (view) {
          const oldSeek = view.SeekView;
          try { view.Type = 3; } catch (e) {}
          try {
            view.SeekView = 6; // primary footer
            probe.seekView6Works = true;
            const sel = app.Selection;
            if (sel && typeof sel.TypeText === "function") {
              try {
                sel.TypeText("\u2060");
                probe.selTypeTextWorks = true;
                // 撤销
                try { sel.WholeStory?.(); sel.Delete?.(); } catch (e) {}
              } catch (e) {}
            }
          } catch (e) {}
          try { view.SeekView = oldSeek; } catch (e) {}
        }
      } catch (e) {}
    return probe;
  }

  async function apply() {
    const rules = await loadRules();
    const doc = getActiveDocument();
    if (!doc) {
      return { ok: false, message: "未检测到 WPS Writer 活动文档，请在文字文档中使用。", changed: 0 };
    }

    const rows = collectParagraphs(doc).filter((x) => x.text);
    if (!rows.length) {
      try {
        const raw = normalizeParagraphText(doc?.Content?.Text || doc?.Range?.Text || "");
        if (raw) {
          return {
            ok: false,
            message: "检测到文档有文本，但当前环境无法读取段落对象。请优先通过 Ribbon 的“公文体检/一键规范化”按钮触发。",
            changed: 0
          };
        }
      } catch (e) {}
      return { ok: false, message: "文档内容为空，无法执行规范化。", changed: 0 };
    }

    let changed = 0;
    changed += tryFormatPage(doc, rules);
    changed += addPageNumber(doc, rules);

    // Apply baseline body layout on full content first, then paragraph-level refinement.
    try {
      if (setRangeBodyLayout(doc.Content, rules)) changed += 1;
    } catch (e) {}

    const titleRow = rows[0];
    if (titleRow) {
      if (setParagraphStyle(titleRow.paragraph, rules.title, { align: 1 })) changed += 1; // center
      try {
        titleRow.paragraph.Range.ParagraphFormat.FirstLineIndent = 0;
      } catch (e) {}
    }

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const level = detectHeadingLevel(row.text, rules);
      if (level) {
        const spec = rules.headings[level];
        if (setParagraphStyle(row.paragraph, spec)) changed += 1;
        try {
          row.paragraph.Range.ParagraphFormat.FirstLineIndent = 0;
        } catch (e) {}
      } else {
        if (setParagraphStyle(row.paragraph, rules.body)) changed += 1;
        if (setBodyParagraphLayout(row.paragraph, rules)) changed += 1;
      }
    }

    const report = await scan();
    let lineSpacingVerified = false;
    try {
      const probe = rows[Math.min(1, rows.length - 1)]?.paragraph?.Range?.ParagraphFormat;
      const v = Number(probe?.LineSpacing || 0);
      lineSpacingVerified = v > 0;
    } catch (e) {}
    const pageNumberVerified = hasPageNumber(doc);
    const pageDiagnostics = collectPageNumberDiagnostics(doc);
    return {
      ok: true,
      standard: rules.standard,
      changed,
      remainingIssues: report.issues || [],
      diagnostics: {
        lineSpacingVerified,
        pageNumberVerified,
        page: pageDiagnostics
      }
    };
  }

  window.GovDocFormatter = {
    loadRules,
    scan,
    apply
  };
})();


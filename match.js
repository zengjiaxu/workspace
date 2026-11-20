(function (root) {
  "use strict";

  var EXPECTED_HEADERS = ["考题", "答案", "选项A", "选项B", "选项C", "选项D"];
  var HEADER_COLUMNS = ["A", "B", "C", "D", "E", "F"];

  function getFflate() {
    if (root.fflate && typeof root.fflate.unzipSync === "function") {
      return root.fflate;
    }
    if (typeof require === "function") {
      return require("./vendor/fflate.umd.js");
    }
    throw new Error("无法读取 Excel。");
  }

  function decodeXml(value) {
    return String(value)
      .replace(/&#x([0-9a-fA-F]+);/g, function (_, hex) {
        return String.fromCodePoint(parseInt(hex, 16));
      })
      .replace(/&#(\d+);/g, function (_, dec) {
        return String.fromCodePoint(parseInt(dec, 10));
      })
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&amp;/g, "&");
  }

  function decodeHtml(value) {
    return decodeXml(String(value).replace(/&nbsp;/gi, " "));
  }

  function attr(tag, name) {
    var match = new RegExp("(?:^|\\s)" + name + '="([^"]*)"').exec(tag);
    return match ? decodeXml(match[1]) : "";
  }

  function stripTags(html) {
    return decodeHtml(
      String(html)
        .replace(/<!--[\s\S]*?-->/g, "")
        .replace(/<[^>]+>/g, "")
    );
  }

  function normalizeText(value) {
    return String(value == null ? "" : value)
      .replace(/\uFEFF/g, "")
      .normalize("NFKC")
      .replace(/[\s\u00a0\u3000\u200b\u200c\u200d]+/g, "")
      .replace(/[\p{P}\p{S}]+/gu, "")
      .toLowerCase();
  }

  function textNodes(xml) {
    var parts = [];
    var re = /<(?:[A-Za-z0-9]+:)?t\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z0-9]+:)?t>/g;
    var match;
    while ((match = re.exec(xml))) {
      parts.push(decodeXml(match[1]));
    }
    return parts.join("");
  }

  function cellValue(attrs, inner, shared) {
    var type = attr(attrs, "t");
    if (type === "inlineStr") {
      return textNodes(inner);
    }
    var valueMatch = /<(?:[A-Za-z0-9]+:)?v\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z0-9]+:)?v>/.exec(
      inner
    );
    var raw = valueMatch ? decodeXml(valueMatch[1]) : "";
    if (type === "s") {
      var sharedText = shared[Number(raw)];
      return sharedText == null ? "" : sharedText;
    }
    return raw;
  }

  function parseSharedStrings(xml) {
    if (!xml) return [];
    var strings = [];
    var re = /<(?:[A-Za-z0-9]+:)?si\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z0-9]+:)?si>/g;
    var match;
    while ((match = re.exec(xml))) {
      strings.push(textNodes(match[1]));
    }
    return strings;
  }

  function parseSheetRows(xml, shared) {
    var rows = [];
    var rowRe = /<(?:[A-Za-z0-9]+:)?row\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z0-9]+:)?row>/g;
    var rowMatch;
    while ((rowMatch = rowRe.exec(xml))) {
      var cells = {};
      var cellRe =
        /<(?:[A-Za-z0-9]+:)?c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:[A-Za-z0-9]+:)?c>)/g;
      var cellMatch;
      while ((cellMatch = cellRe.exec(rowMatch[1]))) {
        var ref = attr(cellMatch[1], "r");
        var column = ((ref.match(/^[A-Z]+/i) || [""])[0] || "").toUpperCase();
        if (!column) continue;
        cells[column] = cellValue(cellMatch[1], cellMatch[2] || "", shared).trim();
      }
      rows.push(cells);
    }
    return rows;
  }

  function decodeZipText(bytes) {
    if (!bytes) return "";
    return new TextDecoder("utf-8").decode(bytes).replace(/^\uFEFF/, "");
  }

  function worksheetPath(files) {
    var rels = decodeZipText(files["xl/_rels/workbook.xml.rels"]);
    var relRe = /<(?:[A-Za-z0-9]+:)?Relationship\b[^>]*>/g;
    var match;
    while ((match = relRe.exec(rels))) {
      if (!/\/worksheet"/.test(attr(match[0], "Type"))) continue;
      var target = attr(match[0], "Target").replace(/^\//, "");
      if (!target) continue;
      if (target.indexOf("xl/") !== 0) target = "xl/" + target;
      if (files[target]) return target;
    }
    var names = Object.keys(files);
    for (var i = 0; i < names.length; i++) {
      if (/^xl\/worksheets\/sheet\d+\.xml$/.test(names[i])) return names[i];
    }
    throw new Error("无法读取这个 Excel。");
  }

  function sameHeaders(row) {
    for (var i = 0; i < EXPECTED_HEADERS.length; i++) {
      var actual = String(row[HEADER_COLUMNS[i]] || "")
        .trim()
        .normalize("NFKC");
      if (actual !== EXPECTED_HEADERS[i]) return false;
    }
    return true;
  }

  function parseXiaobaoXlsx(buffer) {
    var bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
    var files;
    try {
      files = getFflate().unzipSync(bytes);
    } catch (error) {
      throw new Error("无法读取这个 Excel。");
    }
    var shared = parseSharedStrings(decodeZipText(files["xl/sharedStrings.xml"]));
    var sheet = decodeZipText(files[worksheetPath(files)]);
    var rows = parseSheetRows(sheet, shared);
    if (!rows.length || !sameHeaders(rows[0])) {
      throw new Error(
        "这不是小包搜题导入格式。表头需要是：考题、答案、选项A、选项B、选项C、选项D。"
      );
    }
    var bank = [];
    for (var i = 1; i < rows.length; i++) {
      var stem = rows[i].A || "";
      if (!stem) continue;
      var options = [];
      for (var c = 2; c < HEADER_COLUMNS.length; c++) {
        var option = rows[i][HEADER_COLUMNS[c]] || "";
        if (option) options.push(option);
      }
      bank.push({
        stem: stem,
        answer: rows[i].B || "",
        options: options,
      });
    }
    if (!bank.length) {
      throw new Error("题库里没有题目。");
    }
    return bank;
  }

  function cleanStem(raw) {
    var text = String(raw).replace(/\s+/g, " ").trim().normalize("NFKC");
    text = text.replace(/^\[(?:单选|多选|判断)\]\s*/, "");
    var number = "";
    var match = /^(\d+)\s*[.、]\s*/.exec(text);
    if (match) {
      number = match[1];
      text = text.slice(match[0].length);
    }
    text = text.replace(/[(（]\s*\d+(?:\.\d+)?\s*分\s*[)）]\s*$/, "").trim();
    return { number: number, stem: text };
  }

  function parseExamHtml(html) {
    var source = String(html);
    var questionRe =
      /<div\b[^>]*class="[^"]*\bpc_examtest_content_question\b[^"]*"[^>]*>([\s\S]*?)<\/div>/g;
    var found = [];
    var match;
    while ((match = questionRe.exec(source))) {
      found.push({
        html: match[1],
        start: match.index,
        end: questionRe.lastIndex,
      });
    }
    return found.map(function (item, index) {
      var nextStart = index + 1 < found.length ? found[index + 1].start : source.length;
      var region = source.slice(item.end, nextStart);
      var listMatch =
        /<div\b[^>]*class="[^"]*\bpc_examtest_content_list\b[^"]*"[^>]*>([\s\S]*?)<\/div>\s*<div\b[^>]*class="[^"]*\bhidden\b/.exec(
          region
        );
      var options = [];
      if (listMatch) {
        var optionRe =
          /<span\b[^>]*class="[^"]*\bexamtest_content_list_des_optionsort\b[^"]*"[^>]*>[\s\S]*?<\/span>([\s\S]*?)<\/div>/g;
        var optionMatch;
        while ((optionMatch = optionRe.exec(listMatch[1]))) {
          var optionText = stripTags(optionMatch[1]).replace(/\s+/g, " ").trim();
          if (optionText) options.push(optionText);
        }
      }
      var withoutScore = item.html.replace(
        /<span\b[^>]*class="[^"]*\bpc_examtest_content_question_value\b[^"]*"[^>]*>[\s\S]*?<\/span>/g,
        ""
      );
      var cleaned = cleanStem(stripTags(withoutScore));
      return {
        number: cleaned.number || String(index + 1),
        stem: cleaned.stem,
        options: options,
      };
    });
  }

  function optionKey(options) {
    return (options || []).map(normalizeText).join("\n");
  }

  function chooseCandidate(question, candidates) {
    if (candidates.length === 1) return candidates[0];
    var wanted = optionKey(question.options);
    if (wanted) {
      var byOption = candidates.filter(function (item) {
        return optionKey(item.options) === wanted;
      });
      if (byOption.length === 1) return byOption[0];
      if (byOption.length > 1) candidates = byOption;
    }
    var first = candidates[0].answer;
    for (var i = 1; i < candidates.length; i++) {
      if (candidates[i].answer !== first) return null;
    }
    return candidates[0];
  }

  function matchQuestions(questions, bank) {
    var groups = new Map();
    bank.forEach(function (item) {
      var key = normalizeText(item.stem);
      if (!key) return;
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(item);
    });
    var unmatched = 0;
    var rows = questions.map(function (question) {
      var key = normalizeText(question.stem);
      var candidates = key ? groups.get(key) : null;
      var chosen = candidates && candidates.length ? chooseCandidate(question, candidates) : null;
      if (!chosen) {
        unmatched += 1;
        return { number: question.number, answer: "" };
      }
      return { number: question.number, answer: chosen.answer };
    });
    return {
      rows: rows,
      unmatched: unmatched,
      matched: rows.length - unmatched,
    };
  }

  var api = {
    normalizeText: normalizeText,
    parseExamHtml: parseExamHtml,
    parseXiaobaoXlsx: parseXiaobaoXlsx,
    matchQuestions: matchQuestions,
  };

  if (typeof module !== "undefined" && module.exports) {
    module.exports = api;
  }
  root.ExamMatch = api;
})(typeof globalThis !== "undefined" ? globalThis : this);

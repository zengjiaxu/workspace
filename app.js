(function () {
  "use strict";

  var htmlInput = document.querySelector("#exam-html");
  var xlsxInput = document.querySelector("#bank-xlsx");
  var htmlName = document.querySelector("#exam-name");
  var xlsxName = document.querySelector("#bank-name");
  var button = document.querySelector("#run");
  var statusEl = document.querySelector("#status");
  var summary = document.querySelector("#summary");
  var result = document.querySelector("#result");

  function fileLabel(input) {
    return input.files && input.files[0] ? input.files[0].name : "未选择文件";
  }

  function clearResult() {
    statusEl.textContent = "";
    summary.hidden = true;
    summary.textContent = "";
    result.replaceChildren();
  }

  function refresh() {
    htmlName.textContent = fileLabel(htmlInput);
    xlsxName.textContent = fileLabel(xlsxInput);
    button.disabled = !(htmlInput.files[0] && xlsxInput.files[0]);
  }

  function decodeHtmlFile(buffer) {
    var head = new TextDecoder("utf-8").decode(buffer.slice(0, 4096));
    var charsetMatch = /charset\s*=\s*["']?\s*([A-Za-z0-9._-]+)/i.exec(head);
    var charset = charsetMatch ? charsetMatch[1] : "utf-8";
    try {
      return new TextDecoder(charset).decode(buffer);
    } catch (error) {
      return new TextDecoder("utf-8").decode(buffer);
    }
  }

  function render(rows) {
    var table = document.createElement("table");
    var thead = document.createElement("thead");
    var headRow = document.createElement("tr");
    ["序号", "答案"].forEach(function (label) {
      var th = document.createElement("th");
      th.scope = "col";
      th.textContent = label;
      headRow.appendChild(th);
    });
    thead.appendChild(headRow);
    var tbody = document.createElement("tbody");
    rows.forEach(function (row) {
      var tr = document.createElement("tr");
      var numberCell = document.createElement("td");
      var answerCell = document.createElement("td");
      numberCell.textContent = row.number;
      answerCell.textContent = row.answer;
      tr.appendChild(numberCell);
      tr.appendChild(answerCell);
      tbody.appendChild(tr);
    });
    table.appendChild(thead);
    table.appendChild(tbody);
    result.replaceChildren(table);
  }

  htmlInput.addEventListener("change", function () {
    clearResult();
    refresh();
  });
  xlsxInput.addEventListener("change", function () {
    clearResult();
    refresh();
  });

  button.addEventListener("click", function () {
    var htmlFile = htmlInput.files[0];
    var xlsxFile = xlsxInput.files[0];
    if (!htmlFile || !xlsxFile) return;
    clearResult();
    button.disabled = true;
    Promise.all([htmlFile.arrayBuffer(), xlsxFile.arrayBuffer()])
      .then(function (buffers) {
        var questions = ExamMatch.parseExamHtml(decodeHtmlFile(buffers[0]));
        if (!questions.length) {
          throw new Error("没有在 HTML 中找到题目。请选择已保存的考试页面。");
        }
        var bank = ExamMatch.parseXiaobaoXlsx(buffers[1]);
        return ExamMatch.matchQuestions(questions, bank);
      })
      .then(function (matched) {
        summary.hidden = false;
        summary.textContent = "未匹配 " + matched.unmatched + " 题";
        summary.classList.toggle("is-clear", matched.unmatched === 0);
        render(matched.rows);
      })
      .catch(function (error) {
        statusEl.textContent = error && error.message ? error.message : "查询失败。";
      })
      .finally(refresh);
  });

  refresh();
})();

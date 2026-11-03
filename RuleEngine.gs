/** Basic rule engine; accepts plain grid data, no Google services. */
function inspectGrid_(input, limit) {
  const values = input.values || [];
  const formulas = input.formulas || [];
  const height = values.length;
  const width = height ? values[0].length : 0;
  const findings = [];
  let truncated = false;
  const max = Math.max(0, limit);
  function add(id, severity, a1, reason, fix) {
    if (findings.length >= max) {
      truncated = true;
      return;
    }
    findings.push({
      ruleId: id,
      severity: severity,
      sheetName: input.sheetName,
      cell: a1,
      reason: reason,
      suggestion: fix,
      gid: input.gid,
    });
  }
  if (!height || !width) return { findings: findings, truncated: false };

  const occupiedRows = Array(height).fill(0);
  const occupiedCols = Array(width).fill(0);
  const colHasDataBelowHeader = Array(width).fill(false);
  for (let r = 0; r < height; r++) {
    for (let c = 0; c < width; c++) {
      if (hasContent_(values[r][c]) || !!getMatrixCell_(formulas, r, c)) {
        occupiedRows[r]++;
        occupiedCols[c]++;
        if (r > 0) colHasDataBelowHeader[c] = true;
      }
    }
  }
  if (occupiedRows.every(function (count) { return count === 0; })) {
    // Even visually empty sheets may contain off-grid objects or hidden fields.
    inspectAdditionalRules_(input, add);
    return { findings: findings, truncated: truncated };
  }

  // L1-12: merged cells in the used range.
  (input.mergedA1 || []).forEach(function (a1) {
    add('L1-12', '要修正', a1, 'データ範囲に結合セルがあります。', '結合を解除し、各セルを独立した値として扱ってください。');
  });

  // L3-01: clearly blank first row or apparent title row.
  if (height > 1 && occupiedRows[0] === 0 && occupiedRows.slice(1).some(function (n) { return n > 0; })) {
    add('L3-01', '要確認', 'A1', '先頭行が空白です。', '公開・連携用データでは1行目から列見出しを配置してください。');
  } else if (height > 1 && occupiedRows[0] === 1 && occupiedRows[1] >= 2) {
    add('L3-01', '要確認', 'A1:' + columnA1_(width) + '2', '1行目が表題、2行目が列見出しである可能性があります。', '表題は別シートなどで管理し、データは1行目を列見出しにしてください。');
  }

  // L1-05 / L2-03: assumes a single-row header at row 1.
  const headers = new Map();
  for (let c = 0; c < width; c++) {
    if (!occupiedCols[c]) continue;
    const value = values[0][c];
    const header = String(value == null ? '' : value).trim();
    const a1 = columnA1_(c + 1) + '1';
    if (!header) {
      if (colHasDataBelowHeader[c]) {
        add('L1-05', '要修正', a1, 'データがある列に見出しがありません（1行目を見出しと仮定）。', 'その列の意味がわかる固有の列見出しを1行目に入力してください。');
      }
    } else {
      if (headers.has(header)) {
        add('L2-03', '要修正', a1, '列見出し「' + clipText_(header) + '」が ' + headers.get(header) + ' と重複しています。', '列ごとに一意で完結した見出しへ変更してください。');
      } else {
        headers.set(header, a1);
      }
    }
  }

  // L1-03: entire empty lines/columns between non-empty lines/columns.
  // Precompute bounds to avoid repeated O(n) slice scans on large sheets.
  const firstRow = occupiedRows.findIndex(Boolean);
  const lastRow = occupiedRows.length - 1 - occupiedRows.slice().reverse().findIndex(Boolean);
  const firstCol = occupiedCols.findIndex(Boolean);
  const lastCol = occupiedCols.length - 1 - occupiedCols.slice().reverse().findIndex(Boolean);
  for (let r = Math.max(1, firstRow + 1); r < lastRow; r++) {
    if (occupiedRows[r] === 0) {
      add('L1-03', '要確認', String(r + 1) + ':' + String(r + 1), 'データの途中に空白行があります。', '装飾・区切り用の行であれば削除し、1つの連続した表にしてください。');
    }
  }
  for (let c = Math.max(1, firstCol + 1); c < lastCol; c++) {
    if (occupiedCols[c] === 0) {
      const letter = columnA1_(c + 1);
      add('L1-03', '要確認', letter + ':' + letter, 'データの途中に空白列があります。', '装飾・区切り用の列であれば削除し、表を分断しないでください。');
    }
  }

  for (let r = 0; r < height; r++) {
    for (let c = 0; c < width; c++) {
      const a1 = columnA1_(c + 1) + String(r + 1);
      const formula = getMatrixCell_(formulas, r, c);
      const value = values[r][c];
      if (formula) {
        add('L2-06', '要確認', a1, '数式が残っています。', '外部連携・公開用の確定データでは、必要に応じて計算結果を値に置き換えてください。業務用の数式は自動削除しません。');
      }
      if (typeof value !== 'string' || value === '') continue;
      if (/^[ \t\u3000]|[ \t\u3000]$/.test(value)) {
        add('L1-06', '要確認', a1, 'セルの先頭または末尾に空白文字があります。', '意図しない余分な空白・タブを取り除いてください。');
      }
      if (/[\r\n]/.test(value)) {
        add('L1-07', '要確認', a1, 'セル内に改行があり、複数データが含まれる可能性があります。', '一つのセルに複数の独立値を入れていないか確認し、必要なら列や行に分けてください。');
      }
      if (r > 0 && looksLikeNumericText_(value)) {
        add('L2-01', '要確認', a1, '数値に見える値「' + clipText_(value) + '」が文字列で保存されています。', '数値を数値型で保存し、単位や注釈は別列または表示形式に分離してください。コード・識別子なら修正不要です。');
      }
      if (r > 0 && /^その他[（(][^\r\n）)]{1,}[）)]$/.test(value)) {
        add('L2-05', '要確認', a1, '「その他」と具体的な自由記述が同じセルに入っています。', 'カテゴリ列には「その他」のみを入れ、内容は別の詳細列へ分離することを検討してください。');
      }
    }
  }
  inspectAdditionalRules_(input, add);
  return { findings: findings, truncated: truncated };
}

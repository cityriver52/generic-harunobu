/**
 * Deterministic/heuristic additions corresponding to the published Harunobu rules.
 * This code does not invoke an AI API and does not reuse the official Python code.
 * All potential meaning-dependent defects are reported as 要確認.
 */
function inspectAdditionalRules_(input, add) {
  const rows = input.values || [];
  const height = rows.length;
  const width = height ? rows[0].length : 0;
  if (!height || !width) return;
  const text = function(value) { return value == null ? '' : String(value).trim(); };
  const present = function(value) { return value !== '' && value !== null && value !== undefined; };
  const pos = function(r, c) { return columnA1_(c + 1) + (r + 1); };
  const headers = rows[0].map(text);
  const rowCells = rows.map(function(row) { return row.filter(present).length; });
  const normalize = function(value) {
    return text(value).normalize('NFKC').replace(/[\s\u3000]+/g, '').toLowerCase();
  };
  const dateLike = function(value) {
    return /^(?:19|20)\d{2}(?:年|[-/])|^(?:令和|平成|昭和)\d{1,2}年|^[RHS]\d{1,2}(?:年)?$/i.test(text(value));
  };

  // L1-02: repeated header patterns are a deterministic clue of multiple tables.
  if (height >= 4 && width >= 2) {
    const expected = headers.filter(function(v) { return v !== ''; }).length;
    if (expected >= 2) {
      for (let r = 2; r < height; r++) {
        let matches = 0;
        for (let c = 0; c < width; c++) if (headers[c] && text(rows[r][c]) === headers[c]) matches++;
        if (matches >= 2 && matches / expected >= 0.65) {
          add('L1-02', '要確認', pos(r, 0) + ':' + pos(r, width - 1),
            'データの途中に見出し行が再出現しています。別の表を連結している可能性があります。',
            '複数の表なら別シートに分離し、レコードは1つの見出し行の下へ配置してください。');
        }
      }
    }
  }

  // L1-04: detached notes/footers that look unrelated to records.
  if (height > 2 && width > 1) {
    for (let r = 2; r < height; r++) {
      if (rowCells[r] !== 1) continue;
      const joined = rows[r].map(text).filter(Boolean).join(' ');
      if (/^(?:[※＊*]|注[：:]?|備考[：:]?|出典[：:]?|資料[：:]?|最終更新[日：:]?|作成者[：:]?)/.test(joined)) {
        add('L1-04', '要確認', pos(r, 0) + ':' + pos(r, width - 1),
          '表の途中または末尾に注記・出典などの独立した行があります。',
          'データのレコードと無関係な情報なら別シートやメタデータに分離してください。');
      }
    }
  }

  // L1-08: known compatibility and private-use characters (names can be exceptions).
  const unsupported = /[\u2460-\u24ff\u2160-\u2188\u3370-\u33ff\uE000-\uF8FF]/u;
  for (let r = 0; r < height; r++) {
    for (let c = 0; c < width; c++) {
      const value = rows[r][c];
      if (typeof value === 'string' && unsupported.test(value)) {
        add('L1-08', '要確認', pos(r, c),
          '環境依存・互換文字の疑いがある文字が含まれます。',
          '必要なら標準文字へ置換し、固有名詞等で必須の場合は元の表記を維持してください。');
      }
    }
  }

  // L1-09: only potential ambiguity; blanks and zero are different values.
  for (let c = 0; c < width; c++) {
    let zeros = 0, blanks = 0, numeric = 0, example = -1;
    for (let r = 1; r < height; r++) {
      const value = rows[r][c];
      if (typeof value === 'number') { numeric++; if (value === 0) zeros++; }
      if (!present(value) && rowCells[r] > 0) { blanks++; example = r; }
    }
    if (numeric >= 2 && zeros > 0 && blanks > 0) {
      add('L1-09', '要確認', pos(example, c),
        '同じ数値列に「0」と空欄があります。欠損・非該当とゼロの区別が必要か確認してください。',
        '0と欠損値の意味を定義し、欠損コードや説明を別途示してください。空欄が必ず誤りとは限りません。');
    }
  }

  // L1-10: objects, drawings and charts obtained from the Sheets API metadata.
  (input.objects || []).forEach(function(object) {
    add('L1-10', '要確認', object.cell || 'A1',
      'シート上に' + (object.type || 'オブジェクト') + 'があります。重要な情報がセル外にある可能性があります。',
      '必要な注釈・値はセルや別表へ移し、オブジェクトは補助資料として扱ってください。');
  });
  // L1-13: only explicitly sampled hidden rows/columns.
  (input.hiddenRows || []).forEach(function(index) {
    add('L1-13', '要確認', index + ':' + index,
      '使用範囲内の行が非表示になっています。',
      '公開・連携用データに不要であれば行ごと削除し、必要な行なら表示してください。');
  });
  (input.hiddenColumns || []).forEach(function(index) {
    const col = columnA1_(index);
    add('L1-13', '要確認', col + ':' + col,
      '使用範囲内の列が非表示になっています。',
      '公開・連携用データに不要であれば列ごと削除し、必要な列なら表示してください。');
  });

  // L2-02: common group-by columns whose blank cells may stand for "same as above".
  for (let c = 0; c < width; c++) {
    if (!/(都道府県|市区町村|地域|部署|所属|組織|分類|区分|年度|年月|種別)/.test(headers[c])) continue;
    for (let r = 2; r < height; r++) {
      if (!present(rows[r][c]) && present(rows[r - 1][c]) && rowCells[r] > 0) {
        add('L2-02', '要確認', pos(r, c),
          '分類・地域等の列が空欄です。直前の行と同じ値を省略している可能性があります。',
          '各行が単独で意味を持つように値を明記してください。実際の欠損なら埋めないでください。');
      }
    }
  }

  // L2-04: orthographic variation only; semantic synonyms require external context.
  for (let c = 0; c < width; c++) {
    const variants = new Map();
    for (let r = 1; r < height; r++) {
      if (typeof rows[r][c] !== 'string') continue;
      const v = text(rows[r][c]);
      if (!v || v.length > 35) continue;
      const key = normalize(v);
      if (!variants.has(key)) variants.set(key, new Map());
      if (!variants.get(key).has(v)) variants.get(key).set(v, r);
    }
    for (const [normalized, forms] of variants.entries()) {
      if (normalized && forms.size > 1) {
        const example = [...forms.entries()][1];
        add('L2-04', '要確認', pos(example[1], c),
          '同じ列に表記が異なるが正規化すると一致する値があります：' +
          [...forms.keys()].slice(0, 4).join(' / '),
          '表記を統一するか、意図して別の分類としている場合は定義を残してください。');
      }
    }
  }

  // L3-03: unqualified special signs inside otherwise numeric columns.
  for (let c = 0; c < width; c++) {
    let numeric = 0;
    for (let r = 1; r < height; r++) if (typeof rows[r][c] === 'number') numeric++;
    if (numeric < 2) continue;
    for (let r = 1; r < height; r++) {
      if (/^(?:[-―ー－…・×Xx*＊]+|秘匿|非公開|非公表|対象外|N\/A)$/i.test(text(rows[r][c]))) {
        add('L3-03', '要確認', pos(r, c),
          '数値主体の列に特殊記号・秘匿表現があります。',
          'この記号の意味、欠損値との違い、処理方法をメタデータで定義してください。');
      }
    }
  }

  // L3-04: numerical measures without explicit units; headers are heuristic.
  const units = /[（(](?:円|千円|百万円|人|件|個|匹|冊|%|％|kg|㎏|m|km|㎡|秒|分|時間|歳)[）)]|(?:円|人|件|個|匹|冊|％|%|km|kg|㎡)$/i;
  for (let c = 0; c < width; c++) {
    const header = headers[c];
    if (!/(?:金額|売上|収入|支出|費用|重量|距離|面積|人口|人数|件数|割合|比率|率$)/.test(header)) continue;
    if (units.test(header)) continue;
    let numeric = 0;
    for (let r = 1; r < height; r++) if (typeof rows[r][c] === 'number') numeric++;
    if (numeric >= 2) {
      add('L3-04', '要確認', pos(0, c),
        '数値列の見出しから単位が読み取れない可能性があります。',
        '見出しに「金額（円）」等の単位を明記するか、別の定義表で単位を説明してください。');
    }
  }

  // L3-05: Japanese-era abbreviations not accompanied by Gregorian year.
  for (let c = 0; c < width; c++) {
    for (let r = 0; r < height; r++) {
      const value = text(rows[r][c]);
      if (!value || !/(?:令和|平成|昭和)\d{1,2}年|^[RHS]\d{1,2}(?:年度?)?$/i.test(value)) continue;
      if (!/(?:19|20)\d{2}/.test(value)) {
        add('L3-05', '要確認', pos(r, c),
          '和暦のみ、または和暦略記で時点を示しています。',
          '西暦4桁を併記するか、e-Stat等の標準化した時間軸コードを使用してください。');
      }
    }
  }

  // L3-06: unambiguous common prefecture abbreviations in a prefecture column.
  const prefectures = ['北海道','青森県','岩手県','宮城県','秋田県','山形県','福島県',
    '茨城県','栃木県','群馬県','埼玉県','千葉県','東京都','神奈川県','新潟県',
    '富山県','石川県','福井県','山梨県','長野県','岐阜県','静岡県','愛知県',
    '三重県','滋賀県','京都府','大阪府','兵庫県','奈良県','和歌山県',
    '鳥取県','島根県','岡山県','広島県','山口県','徳島県','香川県',
    '愛媛県','高知県','福岡県','佐賀県','長崎県','熊本県','大分県',
    '宮崎県','鹿児島県','沖縄県'];
  const shortened = new Map(prefectures.map(function(full) {
    return [full.slice(0, -1), full];
  }).filter(function(pair) { return pair[0] !== '北海'; }));
  for (let c = 0; c < width; c++) {
    if (!/都道府県|府県名|県名/.test(headers[c])) continue;
    for (let r = 1; r < height; r++) {
      const value = text(rows[r][c]);
      if (shortened.has(value)) {
        add('L3-06', '要確認', pos(r, c),
          '都道府県名が省略形「' + value + '」になっている可能性があります。',
          '正式名称「' + shortened.get(value) + '」または統計用地域コードを検討してください。');
      }
    }
  }

  // L3-09: a time series spread horizontally over three or more columns.
  let runStart = -1;
  for (let c = 0; c <= width; c++) {
    const isPeriod = c < width && dateLike(headers[c]);
    if (isPeriod && runStart === -1) runStart = c;
    if (!isPeriod && runStart !== -1) {
      if (c - runStart >= 3) {
        add('L3-09', '要確認', pos(0, runStart) + ':' + pos(0, c - 1),
          '時点が3列以上に横展開されています。横持ち形式の可能性があります。',
          '年度などを1列の変数とし、観測値を各行に持つ縦持ち形式を検討してください。');
      }
      runStart = -1;
    }
  }
}

/** Lightweight formatting / visibility / object sampling for Google Sheets. */
function sheetInspectionMetadata_(sheet, rows, cols) {
  const info = { hiddenRows: [], hiddenColumns: [], objects: [], limitations: [] };
  const maxRows = Math.min(rows, 120);
  const maxCols = Math.min(cols, 50);
  try {
    if (typeof sheet.isRowHiddenByUser === 'function') {
      for (let r = 1; r <= maxRows; r++) {
        if (sheet.isRowHiddenByUser(r)) info.hiddenRows.push(r);
      }
      if (rows > maxRows) info.limitations.push('L1-13：非表示行は先頭' + maxRows + '行のみ確認');
    } else info.limitations.push('L1-13：非表示行は取得不可');
  } catch (e) { info.limitations.push('L1-13：非表示行の取得失敗'); }
  try {
    if (typeof sheet.isColumnHiddenByUser === 'function') {
      for (let c = 1; c <= maxCols; c++) {
        if (sheet.isColumnHiddenByUser(c)) info.hiddenColumns.push(c);
      }
      if (cols > maxCols) info.limitations.push('L1-13：非表示列は先頭' + maxCols + '列のみ確認');
    } else info.limitations.push('L1-13：非表示列は取得不可');
  } catch (e) { info.limitations.push('L1-13：非表示列の取得失敗'); }

  const sources = [
    { method: 'getImages', type: '画像' },
    { method: 'getDrawings', type: '図形' },
    { method: 'getCharts', type: 'グラフ' }
  ];
  sources.forEach(function(source) {
    if (typeof sheet[source.method] !== 'function') {
      info.limitations.push('L1-10：' + source.type + 'の取得不可');
      return;
    }
    try {
      const objects = sheet[source.method]() || [];
      objects.slice(0, 100).forEach(function(object) {
        let cell = 'A1';
        try {
          if (typeof object.getAnchorCell === 'function' && object.getAnchorCell())
            cell = object.getAnchorCell().getA1Notation();
          else if (typeof object.getContainerInfo === 'function') {
            const anchor = object.getContainerInfo();
            if (anchor && anchor.getAnchorRow && anchor.getAnchorColumn) {
              cell = columnA1_(anchor.getAnchorColumn()) + anchor.getAnchorRow();
            }
          }
        } catch (e) { /* keep safe fallback */ }
        info.objects.push({ type: source.type, cell: cell });
      });
      if (objects.length > 100) info.limitations.push('L1-10：' + source.type + 'は100件まで確認');
    } catch (e) { info.limitations.push('L1-10：' + source.type + 'の取得失敗'); }
  });
  return info;
}

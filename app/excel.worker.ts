/// <reference lib="webworker" />

type Row = Record<string, unknown>;

const aliases = {
  productName: ["상품명", "제품명", "상품이름", "품명", "상품 명"],
  basePrice: ["판매가", "판매가격", "상품가격", "상품 판매가", "공급가", "공급가격", "판매단가", "단가", "기준가격", "원가", "매입가", "소비자가", "정상가", "가격"],
};

function cleanKey(value: unknown) {
  return String(value ?? "").replace(/[\s\n\r_\-()\[\]\/]/g, "").toLowerCase();
}

function normalize(value: unknown) {
  return String(value ?? "").trim();
}

self.onmessage = async (event: MessageEvent<ArrayBuffer>) => {
  try {
    const XLSX = await import("xlsx");
    const workbook = XLSX.read(event.data, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const matrix = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
      header: 1,
      defval: "",
      raw: false,
    });

    const candidates = [...aliases.productName, ...aliases.basePrice].map(cleanKey);
    let headerIndex = 0;
    let bestScore = -1;

    matrix.slice(0, 30).forEach((row, index) => {
      const score = row
        .map(cleanKey)
        .filter(Boolean)
        .filter(key => candidates.some(candidate => key === candidate || key.includes(candidate) || candidate.includes(key))).length;
      if (score > bestScore) {
        bestScore = score;
        headerIndex = index;
      }
    });

    const headers = (matrix[headerIndex] ?? []).map((value, index) => normalize(value) || `열${index + 1}`);
    const rows: Row[] = matrix
      .slice(headerIndex + 1)
      .map(values => Object.fromEntries(headers.map((header, index) => [header, values[index] ?? ""])))
      .filter(row => Object.values(row).some(value => normalize(value)));

    self.postMessage({ ok: true, rows });
  } catch (error) {
    self.postMessage({ ok: false, message: error instanceof Error ? error.message : "파일을 읽지 못했습니다." });
  }
};

export {};

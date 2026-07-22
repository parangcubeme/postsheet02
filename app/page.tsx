"use client";

import { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";

type Row = Record<string, unknown>;
type Market = "스마트스토어" | "옥션" | "G마켓";
type Rule = { market: Market; feeRate: number; marginRate: number; roundUnit: number };

type Product = {
  productName: string;
  sellerCode: string;
  cost: number;
  salePrice: number;
  stock: number;
  optionName: string;
  optionValue: string;
  imageUrl: string;
  detailHtml: string;
  categoryCode: string;
  shippingFee: number;
};

const aliases = {
  productName: ["상품명", "제품명", "상품이름", "품명"],
  sellerCode: ["상품코드", "판매자상품코드", "자체상품코드", "관리코드", "품목코드"],
  cost: ["공급가", "원가", "매입가", "공급가격"],
  stock: ["재고", "재고수량", "수량"],
  optionName: ["옵션명", "옵션항목", "옵션"],
  optionValue: ["옵션값", "옵션내용", "선택옵션"],
  imageUrl: ["대표이미지", "이미지URL", "이미지", "메인이미지"],
  detailHtml: ["상세설명", "상세HTML", "상품상세", "상세페이지"],
  categoryCode: ["카테고리코드", "카테고리번호", "카테고리ID"],
  shippingFee: ["배송비", "기본배송비"],
};

const defaults: Rule[] = [
  { market: "스마트스토어", feeRate: 6, marginRate: 30, roundUnit: 100 },
  { market: "옥션", feeRate: 13, marginRate: 30, roundUnit: 100 },
  { market: "G마켓", feeRate: 13, marginRate: 30, roundUnit: 100 },
];

function normalize(value: unknown) {
  return String(value ?? "").trim();
}

function numberValue(value: unknown) {
  const n = Number(String(value ?? "0").replace(/[^0-9.-]/g, ""));
  return Number.isFinite(n) ? n : 0;
}

function pick(row: Row, names: string[]) {
  const keys = Object.keys(row);
  const key = keys.find((k) => names.some((name) => k.replace(/\s/g, "").includes(name.replace(/\s/g, ""))));
  return key ? row[key] : "";
}

function calculatePrice(cost: number, feeRate: number, marginRate: number, extraCost: number, roundUnit: number) {
  const denominator = 1 - feeRate / 100 - marginRate / 100;
  if (denominator <= 0) return 0;
  const raw = (cost + extraCost) / denominator;
  return Math.ceil(raw / roundUnit) * roundUnit;
}

function toProducts(rows: Row[], feeRate: number, marginRate: number, extraCost: number, roundUnit: number): Product[] {
  return rows
    .map((row) => {
      const cost = numberValue(pick(row, aliases.cost));
      return {
        productName: normalize(pick(row, aliases.productName)),
        sellerCode: normalize(pick(row, aliases.sellerCode)),
        cost,
        salePrice: calculatePrice(cost, feeRate, marginRate, extraCost, roundUnit),
        stock: numberValue(pick(row, aliases.stock)),
        optionName: normalize(pick(row, aliases.optionName)),
        optionValue: normalize(pick(row, aliases.optionValue)),
        imageUrl: normalize(pick(row, aliases.imageUrl)),
        detailHtml: normalize(pick(row, aliases.detailHtml)),
        categoryCode: normalize(pick(row, aliases.categoryCode)),
        shippingFee: numberValue(pick(row, aliases.shippingFee)),
      };
    })
    .filter((row) => row.productName || row.sellerCode);
}

function marketRows(products: Product[], market: Market) {
  if (market === "스마트스토어") {
    return products.map((p) => ({
      상품명: p.productName,
      판매자상품코드: p.sellerCode,
      판매가: p.salePrice,
      재고수량: p.stock,
      카테고리ID: p.categoryCode,
      옵션명: p.optionName,
      옵션값: p.optionValue,
      대표이미지URL: p.imageUrl,
      상세설명: p.detailHtml,
      배송비: p.shippingFee,
    }));
  }

  return products.map((p) => ({
    상품명: p.productName,
    판매자관리코드: p.sellerCode,
    판매가격: p.salePrice,
    수량: p.stock,
    카테고리번호: p.categoryCode,
    주문선택사항명: p.optionName,
    주문선택사항값: p.optionValue,
    기본이미지: p.imageUrl,
    상품상세설명: p.detailHtml,
    배송비: p.shippingFee,
    판매사이트: market,
  }));
}

export default function Home() {
  const [market, setMarket] = useState<Market>("스마트스토어");
  const [rules, setRules] = useState<Rule[]>(defaults);
  const [feeRate, setFeeRate] = useState(6);
  const [marginRate, setMarginRate] = useState(30);
  const [extraCost, setExtraCost] = useState(0);
  const [roundUnit, setRoundUnit] = useState(100);
  const [sourceRows, setSourceRows] = useState<Row[]>([]);
  const [fileName, setFileName] = useState("");
  const [status, setStatus] = useState("샵온 상품 엑셀을 업로드해 주세요.");

  useEffect(() => {
    fetch("/api/rules")
      .then((r) => r.json())
      .then((data) => {
        if (Array.isArray(data.rules) && data.rules.length) setRules(data.rules);
      })
      .catch(() => setRules(defaults));
  }, []);

  useEffect(() => {
    const rule = rules.find((r) => r.market === market) ?? defaults[0];
    setFeeRate(rule.feeRate);
    setMarginRate(rule.marginRate);
    setRoundUnit(rule.roundUnit);
  }, [market, rules]);

  const products = useMemo(
    () => toProducts(sourceRows, feeRate, marginRate, extraCost, roundUnit),
    [sourceRows, feeRate, marginRate, extraCost, roundUnit]
  );

  async function readFile(file?: File) {
    if (!file) return;
    setFileName(file.name);
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json<Row>(sheet, { defval: "" });
      setSourceRows(rows);
      setStatus(`${rows.length}행을 읽었습니다. ${products.length || rows.length}개 상품을 변환할 수 있습니다.`);
    } catch {
      setSourceRows([]);
      setStatus("파일을 읽지 못했습니다. xlsx, xls 또는 csv 파일인지 확인해 주세요.");
    }
  }

  function download() {
    if (!products.length) return;
    const rows = marketRows(products, market);
    const sheet = XLSX.utils.json_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, sheet, market);
    XLSX.writeFile(workbook, `postsheet02_${market}_${new Date().toISOString().slice(0, 10)}.xlsx`);
  }

  return (
    <main className="container">
      <section className="hero">
        <span className="badge">postsheet02 · 상품 대량등록 변환</span>
        <h1>샵온 상품 엑셀을<br />오픈마켓 등록 파일로 변환</h1>
        <p>수수료율과 목표 마진율을 일괄 적용한 뒤 스마트스토어, 옥션, G마켓용 XLSX 파일을 생성합니다.</p>
        <div className="privacy">원본·결과 파일 서버 저장 없음 · 변환 로그 없음 · Google Sheets는 규칙 DB로만 사용</div>
      </section>

      <section className="panel">
        <div className="field full">
          <label>1. 샵온 상품 엑셀 업로드</label>
          <input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => readFile(e.target.files?.[0])} />
          <small>{fileName || "선택된 파일 없음"}</small>
        </div>

        <div className="field">
          <label>2. 출력 마켓</label>
          <select value={market} onChange={(e) => setMarket(e.target.value as Market)}>
            <option>스마트스토어</option>
            <option>옥션</option>
            <option>G마켓</option>
          </select>
        </div>
        <div className="field">
          <label>수수료율 (%)</label>
          <input type="number" value={feeRate} onChange={(e) => setFeeRate(Number(e.target.value))} />
        </div>
        <div className="field">
          <label>목표 마진율 (%)</label>
          <input type="number" value={marginRate} onChange={(e) => setMarginRate(Number(e.target.value))} />
        </div>
        <div className="field">
          <label>상품당 추가비용 (원)</label>
          <input type="number" value={extraCost} onChange={(e) => setExtraCost(Number(e.target.value))} />
        </div>
        <div className="field">
          <label>판매가 올림 단위</label>
          <select value={roundUnit} onChange={(e) => setRoundUnit(Number(e.target.value))}>
            <option value={10}>10원</option>
            <option value={100}>100원</option>
            <option value={500}>500원</option>
            <option value={1000}>1,000원</option>
          </select>
        </div>

        <div className="actions full">
          <button onClick={download} disabled={!products.length}>변환 XLSX 다운로드</button>
          <button className="secondary" onClick={() => { setSourceRows([]); setFileName(""); setStatus("화면 데이터를 비웠습니다."); }}>초기화</button>
        </div>
        <div className="status full">{status}</div>
      </section>

      <section className="preview">
        <div className="previewHead">
          <h2>계산 미리보기</h2>
          <span>{products.length}개 상품</span>
        </div>
        <div className="tableWrap">
          <table>
            <thead><tr><th>상품명</th><th>상품코드</th><th>원가</th><th>계산 판매가</th><th>재고</th></tr></thead>
            <tbody>
              {products.slice(0, 30).map((p, i) => (
                <tr key={`${p.sellerCode}-${i}`}><td>{p.productName}</td><td>{p.sellerCode}</td><td>{p.cost.toLocaleString()}</td><td>{p.salePrice.toLocaleString()}</td><td>{p.stock}</td></tr>
              ))}
              {!products.length && <tr><td colSpan={5} className="empty">엑셀을 업로드하면 이곳에 결과가 표시됩니다.</td></tr>}
            </tbody>
          </table>
        </div>
      </section>
    </main>
  );
}

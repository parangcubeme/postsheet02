"use client";

import { useMemo, useState } from "react";
import * as XLSX from "xlsx";

type Row = Record<string, unknown>;
type AppliedSettings = {
  feeRate: number;
  marginRate: number;
  extraCost: number;
  roundUnit: number;
};

type Product = {
  productName: string;
  sellerCode: string;
  basePrice: number;
  salePrice: number;
  stock: number;
  optionName: string;
  optionValue: string;
  imageUrl: string;
  detailHtml: string;
  categoryCode: string;
  shippingFee: number;
  brand: string;
  maker: string;
  originCode: string;
};

const aliases = {
  productName: ["상품명", "제품명", "상품이름", "품명"],
  sellerCode: ["판매자 상품코드", "판매자상품코드", "상품코드", "자체상품코드", "관리코드", "품목코드"],
  basePrice: ["판매가", "상품가격", "판매가격", "기준가격", "공급가", "원가", "매입가", "공급가격"],
  stock: ["재고수량", "재고", "수량"],
  optionName: ["옵션명", "옵션항목", "옵션"],
  optionValue: ["옵션값", "옵션내용", "선택옵션"],
  imageUrl: ["대표이미지", "대표이미지URL", "이미지URL", "이미지", "메인이미지"],
  detailHtml: ["상세설명", "상세HTML", "상품상세", "상세페이지"],
  categoryCode: ["카테고리코드", "카테고리번호", "카테고리ID"],
  shippingFee: ["기본배송비", "배송비"],
  brand: ["브랜드"],
  maker: ["제조사", "제조자"],
  originCode: ["원산지코드", "원산지"],
};

const naverHeaders = [
  "판매자 상품코드","카테고리코드","상품명","상품상태","판매가","단위가격 사용여부","표시용량","표시단위","총용량","부가세","관부가세","재고수량","옵션형태","옵션명","옵션값","옵션가","옵션 재고수량","직접입력 옵션","추가상품명","추가상품값","추가상품가","추가상품 재고수량","대표이미지","추가이미지","상세설명","브랜드","제조사","제조일자","유효일자","원산지코드","수입사","복수원산지여부","원산지 직접입력","미성년자 구매","배송비 템플릿코드","배송방법","택배사코드","배송비유형","기본배송비","배송비 결제방식","조건부무료-\n상품판매가 합계","수량별부과-수량","구간별-\n2구간수량","구간별-\n3구간수량","구간별-\n3구간배송비","구간별-\n추가배송비","반품배송비","교환배송비","지역별 차등 배송비","별도설치비","상품정보제공고시 템플릿코드","상품정보제공고시\n품명","상품정보제공고시\n모델명","상품정보제공고시\n인증허가사항","상품정보제공고시\n제조자","A/S 템플릿코드","A/S 전화번호","A/S 안내","판매자특이사항","즉시할인 값\n(기본할인)","즉시할인 단위\n(기본할인)","모바일\n즉시할인 값","모바일\n즉시할인 단위","복수구매할인\n조건 값","복수구매할인\n조건 단위","복수구매할인\n값","복수구매할인\n단위","상품구매시 포인트\n지급 값","상품구매시 포인트\n지급 단위","텍스트리뷰 작성시\n지급 포인트","포토/동영상 리뷰 작성시\n지급 포인트","한달사용 텍스트리뷰\n작성시 지급 포인트","한달사용\n포토/동영상리뷰 작성시 지급 포인트","알림받기동의 고객 리뷰 작성 시 지급 포인트","무이자\n할부 개월","사은품","판매자바코드","구매평 노출여부","구매평\n비노출사유","알림받기 동의 고객 전용 여부","ISBN","ISSN","독립출판","출간일","출판사","글작가","그림작가","번역자명","문화비 소득공제","사이즈\n상품군","사이즈\n사이즈명","사이즈\n상세 사이즈","사이즈 \n모델명"
];

const requiredRow = naverHeaders.map((header) => {
  const required = ["카테고리코드","상품명","판매가","재고수량","대표이미지","상세설명","원산지코드"];
  return required.includes(header) ? "필수" : "비필수";
});

const groupRow = naverHeaders.map((_, index) => {
  if (index === 0) return "상품 기본정보";
  if (index === 25) return "상품 주요정보";
  if (index === 34) return "배송정보";
  if (index === 50) return "상품정보제공고시";
  if (index === 55) return "A/S, 특이사항";
  if (index === 59) return "할인/혜택정보";
  if (index === 77) return "기타 정보";
  return "";
});

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

function calculatePrice(basePrice: number, feeRate: number, marginRate: number, extraCost: number, roundUnit: number) {
  if (basePrice <= 0) return 0;
  const raw = (basePrice + extraCost) * (1 + marginRate / 100 + feeRate / 100);
  return Math.ceil(raw / roundUnit) * roundUnit;
}

function toProducts(rows: Row[], settings: AppliedSettings): Product[] {
  return rows.map((row) => {
    const basePrice = numberValue(pick(row, aliases.basePrice));
    return {
      productName: normalize(pick(row, aliases.productName)),
      sellerCode: normalize(pick(row, aliases.sellerCode)),
      basePrice,
      salePrice: calculatePrice(basePrice, settings.feeRate, settings.marginRate, settings.extraCost, settings.roundUnit),
      stock: numberValue(pick(row, aliases.stock)) || 1,
      optionName: normalize(pick(row, aliases.optionName)),
      optionValue: normalize(pick(row, aliases.optionValue)),
      imageUrl: normalize(pick(row, aliases.imageUrl)),
      detailHtml: normalize(pick(row, aliases.detailHtml)),
      categoryCode: normalize(pick(row, aliases.categoryCode)),
      shippingFee: numberValue(pick(row, aliases.shippingFee)),
      brand: normalize(pick(row, aliases.brand)),
      maker: normalize(pick(row, aliases.maker)),
      originCode: normalize(pick(row, aliases.originCode)),
    };
  }).filter((row) => row.productName || row.sellerCode);
}

function productToNaverRow(product: Product) {
  const row = Array(naverHeaders.length).fill("");
  const set = (header: string, value: string | number) => {
    const index = naverHeaders.indexOf(header);
    if (index >= 0) row[index] = value;
  };

  set("판매자 상품코드", product.sellerCode);
  set("카테고리코드", product.categoryCode);
  set("상품명", product.productName);
  set("상품상태", "신상품");
  set("판매가", product.salePrice);
  set("부가세", "과세상품");
  set("재고수량", product.stock);
  if (product.optionName && product.optionValue) {
    set("옵션형태", product.optionName.includes("\n") ? "조합형" : "단독형");
    set("옵션명", product.optionName);
    set("옵션값", product.optionValue);
  }
  set("대표이미지", product.imageUrl);
  set("상세설명", product.detailHtml);
  set("브랜드", product.brand);
  set("제조사", product.maker);
  set("원산지코드", product.originCode);
  set("복수원산지여부", "N");
  set("미성년자 구매", "Y");
  set("배송방법", "택배, 소포, 등기");
  set("배송비유형", product.shippingFee > 0 ? "유료" : "무료");
  set("기본배송비", product.shippingFee);
  if (product.shippingFee > 0) set("배송비 결제방식", "선결제");
  set("반품배송비", product.shippingFee || 3000);
  set("교환배송비", (product.shippingFee || 3000) * 2);
  set("별도설치비", "N");
  set("구매평 노출여부", "Y");
  set("알림받기 동의 고객 전용 여부", "N");
  return row;
}

export default function Home() {
  const [feeRate, setFeeRate] = useState(6);
  const [marginRate, setMarginRate] = useState(30);
  const [extraCost, setExtraCost] = useState(0);
  const [roundUnit, setRoundUnit] = useState(100);
  const [appliedSettings, setAppliedSettings] = useState<AppliedSettings | null>(null);
  const [sourceRows, setSourceRows] = useState<Row[]>([]);
  const [fileName, setFileName] = useState("");
  const [status, setStatus] = useState("상품 일괄목록 엑셀을 업로드해 주세요.");

  const products = useMemo(
    () => appliedSettings ? toProducts(sourceRows, appliedSettings) : [],
    [sourceRows, appliedSettings]
  );

  async function readFile(file?: File) {
    if (!file) return;
    setFileName(file.name);
    setAppliedSettings(null);
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json<Row>(sheet, { defval: "" });
      setSourceRows(rows);
      setStatus(`${rows.length}행을 읽었습니다. 수수료율과 마진율을 입력한 뒤 '위 내용 적용하기'를 눌러 주세요.`);
    } catch {
      setSourceRows([]);
      setStatus("파일을 읽지 못했습니다. xlsx, xls 또는 csv 파일인지 확인해 주세요.");
    }
  }

  function applySettings() {
    if (!sourceRows.length || feeRate < 0 || marginRate < 0) return;
    const next = { feeRate, marginRate, extraCost, roundUnit };
    setAppliedSettings(next);
    setStatus(`적용 완료: 수수료 ${feeRate}%, 추가 마진 ${marginRate}%, ${roundUnit.toLocaleString()}원 단위 올림이 전체 상품에 반영됐습니다.`);
  }

  function downloadNaver() {
    if (!products.length || !appliedSettings) return;
    const dataRows = products.map(productToNaverRow);
    const aoa = [groupRow, naverHeaders, requiredRow, ...dataRows];
    const sheet = XLSX.utils.aoa_to_sheet(aoa);
    sheet["!cols"] = naverHeaders.map((header) => ({ wch: Math.min(Math.max(header.replace(/\n/g, "").length + 3, 12), 28) }));
    sheet["!rows"] = [{ hpt: 24 }, { hpt: 46 }, { hpt: 22 }];
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, sheet, "일괄등록");
    XLSX.writeFile(workbook, `postsheet02_네이버스마트스토어_${new Date().toISOString().slice(0, 10)}.xlsx`);
  }

  function resetAll() {
    setSourceRows([]);
    setFileName("");
    setAppliedSettings(null);
    setStatus("화면 데이터를 비웠습니다.");
  }

  const invalidRates = feeRate < 0 || marginRate < 0;
  const settingsChanged = !!appliedSettings && (
    appliedSettings.feeRate !== feeRate ||
    appliedSettings.marginRate !== marginRate ||
    appliedSettings.extraCost !== extraCost ||
    appliedSettings.roundUnit !== roundUnit
  );

  return (
    <main className="container">
      <section className="hero">
        <span className="badge">postsheet02 · 상품 대량등록 변환</span>
        <h1>상품 일괄목록을<br />네이버 등록 파일로 변환</h1>
        <p>원래 판매가격에 사용자가 입력한 수수료율과 마진율을 더한 뒤, 적용 버튼을 눌러 확정한 가격으로 네이버 파일을 생성합니다.</p>
        <div className="privacy">원본·결과 파일 서버 저장 없음 · 변환 로그 없음 · 브라우저 안에서만 계산 및 다운로드</div>
      </section>

      <section className="panel">
        <div className="field full">
          <label>1. 상품 일괄목록 엑셀 업로드</label>
          <input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => readFile(e.target.files?.[0])} />
          <small>{fileName || "선택된 파일 없음"}</small>
        </div>

        <div className="field">
          <label>2. 네이버 수수료율 (%)</label>
          <input min={0} step={0.1} type="number" value={feeRate} onChange={(e) => setFeeRate(Number(e.target.value))} />
          <small>원래 판매가격에 입력한 수수료율만큼 더합니다.</small>
        </div>
        <div className="field">
          <label>3. 추가 마진율 (%)</label>
          <input min={0} step={0.1} type="number" value={marginRate} onChange={(e) => setMarginRate(Number(e.target.value))} />
          <small>원래 판매가격에 입력한 마진율만큼 더합니다.</small>
        </div>
        <div className="field">
          <label>상품당 추가비용 (원)</label>
          <input min={0} type="number" value={extraCost} onChange={(e) => setExtraCost(Number(e.target.value))} />
        </div>
        <div className="field">
          <label>판매가 올림 단위</label>
          <select value={roundUnit} onChange={(e) => setRoundUnit(Number(e.target.value))}>
            <option value={10}>10원 올림</option>
            <option value={100}>100원 올림</option>
            <option value={500}>500원 올림</option>
            <option value={1000}>1,000원 올림</option>
          </select>
        </div>

        {invalidRates && <div className="status full">수수료율과 마진율은 0 이상이어야 합니다.</div>}
        {settingsChanged && <div className="status full">입력값이 변경됐습니다. 다시 '위 내용 적용하기'를 눌러야 다운로드 파일에 반영됩니다.</div>}

        <div className="actions full">
          <button onClick={downloadNaver} disabled={!products.length || invalidRates || settingsChanged}>네이버 스마트스토어 파일 다운로드</button>
          <button onClick={applySettings} disabled={!sourceRows.length || invalidRates}>위 내용 적용하기</button>
          <button className="secondary" onClick={resetAll}>초기화</button>
        </div>
        <div className="status full">{status}</div>
      </section>

      <section className="preview">
        <div className="previewHead">
          <h2>가격 계산 미리보기</h2>
          <span>{products.length}개 상품</span>
        </div>
        <div className="tableWrap">
          <table>
            <thead><tr><th>상품명</th><th>상품코드</th><th>원래 판매가격</th><th>적용 수수료율</th><th>적용 마진율</th><th>최종 판매가</th></tr></thead>
            <tbody>
              {products.slice(0, 30).map((p, i) => (
                <tr key={`${p.sellerCode}-${i}`}><td>{p.productName}</td><td>{p.sellerCode}</td><td>{p.basePrice.toLocaleString()}</td><td>{appliedSettings?.feeRate}%</td><td>{appliedSettings?.marginRate}%</td><td>{p.salePrice.toLocaleString()}</td></tr>
              ))}
              {!products.length && <tr><td colSpan={6} className="empty">엑셀 업로드 후 '위 내용 적용하기'를 누르면 계산 결과가 표시됩니다.</td></tr>}
            </tbody>
          </table>
        </div>
      </section>
    </main>
  );
}

"use client";

import { useMemo, useState } from "react";
import * as XLSX from "xlsx";

type Row = Record<string, unknown>;
type MarketTab = "smartstore" | "auction" | "gmarket";
type AppliedSettings = {
  feeRate: number;
  marginRate: number;
  extraCost: number;
  roundUnit: number;
  categoryCode: string;
  courierCode: string;
  asPhone: string;
  multipleOrigins: "N" | "Y";
};
type Product = {
  raw: Row;
  productName: string;
  sellerCode: string;
  basePrice: number;
  salePrice: number;
  stock: number;
  optionName: string;
  optionValue: string;
  mainImage: string;
  additionalImages: string;
  detailHtml: string;
  shippingFee: number;
  brand: string;
  maker: string;
  vatType: string;
  originCode: string;
  originDirect: string;
};

const aliases = {
  productName: ["상품명", "제품명", "상품이름", "품명", "상품 명"],
  sellerCode: ["판매자 상품코드", "판매자상품코드", "상품코드", "자체상품코드", "관리코드", "품목코드", "판매자코드"],
  basePrice: ["판매가", "판매가격", "상품가격", "상품 판매가", "소비자가", "정상가", "공급가", "공급가격", "판매단가", "단가", "기준가격", "원가", "매입가", "가격"],
  stock: ["재고수량", "재고", "수량", "판매가능수량"],
  optionName: ["옵션명", "옵션항목", "옵션"],
  optionValue: ["옵션값", "옵션내용", "선택옵션"],
  mainImage: ["대표이미지", "대표이미지URL", "이미지1", "메인이미지", "이미지URL"],
  detailHtml: ["상세설명", "상품설명", "상세HTML", "상품상세", "상세페이지", "상세이미지", "상세이미미"],
  shippingFee: ["기본배송비", "배송비"],
  brand: ["브랜드"],
  maker: ["제조사", "제조자"],
  vatType: ["부가세", "과세구분", "부가세유형"],
  originCode: ["원산지코드"],
  originDirect: ["원산지 직접입력", "원산지직접입력", "원산지", "원산지명"],
};

const naverHeaders = [
  "판매자 상품코드","카테고리코드","상품명","상품상태","판매가","단위가격 사용여부","표시용량","표시단위","총용량","부가세","관부가세","재고수량","옵션형태","옵션명","옵션값","옵션가","옵션 재고수량","직접입력 옵션","추가상품명","추가상품값","추가상품가","추가상품 재고수량","대표이미지","추가이미지","상세설명","브랜드","제조사","제조일자","유효일자","원산지코드","수입사","복수원산지여부","원산지 직접입력","미성년자 구매","배송비 템플릿코드","배송방법","택배사코드","배송비유형","기본배송비","배송비 결제방식","조건부무료-\n상품판매가 합계","수량별부과-수량","구간별-\n2구간수량","구간별-\n3구간수량","구간별-\n3구간배송비","구간별-\n추가배송비","반품배송비","교환배송비","지역별 차등 배송비","별도설치비","상품정보제공고시 템플릿코드","상품정보제공고시\n품명","상품정보제공고시\n모델명","상품정보제공고시\n인증허가사항","상품정보제공고시\n제조자","A/S 템플릿코드","A/S 전화번호","A/S 안내","판매자특이사항","즉시할인 값\n(기본할인)","즉시할인 단위\n(기본할인)","모바일\n즉시할인 값","모바일\n즉시할인 단위","복수구매할인\n조건 값","복수구매할인\n조건 단위","복수구매할인\n값","복수구매할인\n단위","상품구매시 포인트\n지급 값","상품구매시 포인트\n지급 단위","텍스트리뷰 작성시\n지급 포인트","포토/동영상 리뷰 작성시\n지급 포인트","한달사용 텍스트리뷰\n작성시 지급 포인트","한달사용\n포토/동영상리뷰 작성시 지급 포인트","알림받기동의 고객 리뷰 작성 시 지급 포인트","무이자\n할부 개월","사은품","판매자바코드","구매평 노출여부","구매평\n비노출사유","알림받기 동의 고객 전용 여부","ISBN","ISSN","독립출판","출간일","출판사","글작가","그림작가","번역자명","문화비 소득공제","사이즈\n상품군","사이즈\n사이즈명","사이즈\n상세 사이즈","사이즈\n모델명","판매준수가"
];

const groupRow = naverHeaders.map((_, index) =>
  index === 0 ? "상품 기본정보" :
  index === 25 ? "상품 주요정보" :
  index === 34 ? "배송정보" :
  index === 50 ? "상품정보제공고시" :
  index === 55 ? "A/S, 특이사항" :
  index === 59 ? "할인/혜택정보" :
  index === 77 ? "기타 정보" : ""
);

function cleanKey(value: unknown) {
  return String(value ?? "").replace(/[\s\n\r_\-()\[\]]/g, "").toLowerCase();
}
function normalize(value: unknown) { return String(value ?? "").trim(); }
function numberValue(value: unknown) {
  const n = Number(String(value ?? "").replace(/,/g, "").replace(/[^0-9.-]/g, ""));
  return Number.isFinite(n) ? n : 0;
}
function pickExact(row: Row, name: string) {
  const target = cleanKey(name);
  const found = Object.entries(row).find(([key]) => cleanKey(key) === target);
  return found ? found[1] : "";
}
function pick(row: Row, names: string[]) {
  for (const name of names) {
    const value = pickExact(row, name);
    if (normalize(value) !== "") return value;
  }
  const entries = Object.entries(row);
  for (const name of names) {
    const target = cleanKey(name);
    const found = entries.find(([key]) => cleanKey(key).includes(target) || target.includes(cleanKey(key)));
    if (found) return found[1];
  }
  return "";
}
function collectAdditionalImages(row: Row) {
  const direct = normalize(pickExact(row, "추가이미지"));
  if (direct) return direct;
  return Object.entries(row)
    .map(([key, value]) => ({ key: cleanKey(key), value: normalize(value) }))
    .filter(({ key, value }) => /^이미지([2-9]|[1-9][0-9]+)$/.test(key) && value)
    .sort((a, b) => Number(a.key.replace("이미지", "")) - Number(b.key.replace("이미지", "")))
    .map(({ value }) => value)
    .join("\n");
}
function parseSheet(sheet: XLSX.WorkSheet) {
  const matrix = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "", raw: false });
  const candidates = [...aliases.productName, ...aliases.basePrice, ...aliases.sellerCode, ...aliases.stock].map(cleanKey);
  let bestIndex = 0;
  let bestScore = -1;
  matrix.slice(0, 30).forEach((row, index) => {
    const score = row.map(cleanKey).filter(Boolean).reduce((sum, key) =>
      sum + (candidates.some((candidate) => key === candidate || key.includes(candidate) || candidate.includes(key)) ? 1 : 0), 0);
    if (score > bestScore) { bestScore = score; bestIndex = index; }
  });
  const rawHeaders = (matrix[bestIndex] ?? []).map((value, index) => normalize(value) || `열${index + 1}`);
  const seen = new Map<string, number>();
  const headers = rawHeaders.map((header) => {
    const count = (seen.get(header) || 0) + 1;
    seen.set(header, count);
    return count === 1 ? header : `${header}_${count}`;
  });
  const rows = matrix.slice(bestIndex + 1)
    .map((values) => Object.fromEntries(headers.map((header, index) => [header, values[index] ?? ""])))
    .filter((row) => Object.values(row).some((value) => normalize(value) !== ""));
  return { rows, headerRow: bestIndex + 1 };
}
function calculatePrice(basePrice: number, settings: AppliedSettings) {
  if (basePrice <= 0) return 0;
  const raw = (basePrice + settings.extraCost) * (1 + settings.marginRate / 100 + settings.feeRate / 100);
  return Math.ceil(raw / settings.roundUnit) * settings.roundUnit;
}
function toProducts(rows: Row[], settings: AppliedSettings): Product[] {
  return rows.map((raw) => {
    const basePrice = numberValue(pick(raw, aliases.basePrice));
    return {
      raw,
      productName: normalize(pick(raw, aliases.productName)),
      sellerCode: normalize(pick(raw, aliases.sellerCode)),
      basePrice,
      salePrice: calculatePrice(basePrice, settings),
      stock: numberValue(pick(raw, aliases.stock)) || 1,
      optionName: normalize(pick(raw, aliases.optionName)),
      optionValue: normalize(pick(raw, aliases.optionValue)),
      mainImage: normalize(pick(raw, aliases.mainImage)),
      additionalImages: collectAdditionalImages(raw),
      detailHtml: normalize(pick(raw, aliases.detailHtml)),
      shippingFee: numberValue(pick(raw, aliases.shippingFee)),
      brand: normalize(pick(raw, aliases.brand)),
      maker: normalize(pick(raw, aliases.maker)),
      vatType: normalize(pick(raw, aliases.vatType)),
      originCode: normalize(pick(raw, aliases.originCode)),
      originDirect: normalize(pick(raw, aliases.originDirect)),
    };
  }).filter((product) => product.productName || product.sellerCode);
}
function productToNaverRow(product: Product, settings: AppliedSettings) {
  const row = naverHeaders.map((header) => pickExact(product.raw, header));
  const set = (header: string, value: string | number) => {
    const index = naverHeaders.indexOf(header);
    if (index >= 0) row[index] = value;
  };
  const existing = (header: string) => normalize(pickExact(product.raw, header));

  set("판매자 상품코드", product.sellerCode);
  set("카테고리코드", settings.categoryCode);
  set("상품명", product.productName);
  set("상품상태", existing("상품상태") || "신상품");
  set("판매가", product.salePrice);
  set("단위가격 사용여부", existing("단위가격 사용여부") || "N");
  set("부가세", product.vatType || "과세상품");
  set("재고수량", product.stock);

  if (!existing("옵션형태") && product.optionName && product.optionValue) set("옵션형태", "조합형");
  if (!existing("옵션명") && product.optionName) set("옵션명", product.optionName);
  if (!existing("옵션값") && product.optionValue) set("옵션값", product.optionValue);

  set("대표이미지", product.mainImage);
  set("추가이미지", product.additionalImages);
  set("상세설명", product.detailHtml);
  if (!existing("브랜드")) set("브랜드", product.brand);
  if (!existing("제조사")) set("제조사", product.maker);
  set("원산지코드", product.originCode);
  set("복수원산지여부", settings.multipleOrigins);
  set("원산지 직접입력", product.originDirect);
  set("미성년자 구매", existing("미성년자 구매") || "Y");
  set("배송방법", "택배, 소포, 등기");
  set("택배사코드", settings.courierCode);
  set("배송비유형", existing("배송비유형") || (product.shippingFee > 0 ? "유료" : "무료"));
  set("기본배송비", product.shippingFee);
  if (product.shippingFee > 0 && !existing("배송비 결제방식")) set("배송비 결제방식", "선결제");
  if (!existing("반품배송비")) set("반품배송비", product.shippingFee || 3000);
  if (!existing("교환배송비")) set("교환배송비", (product.shippingFee || 3000) * 2);
  set("별도설치비", existing("별도설치비") || "N");
  set("A/S 전화번호", settings.asPhone);
  set("A/S 안내", existing("A/S 안내") || "판매자에 문의하시거나, A/S연락처로 문의 주시기 바랍니다.");
  set("구매평 노출여부", existing("구매평 노출여부") || "Y");
  set("알림받기 동의 고객 전용 여부", existing("알림받기 동의 고객 전용 여부") || "N");
  return row;
}

export default function Home() {
  const [activeMarket, setActiveMarket] = useState<MarketTab>("smartstore");
  const [feeRate, setFeeRate] = useState(6);
  const [marginRate, setMarginRate] = useState(30);
  const [extraCost, setExtraCost] = useState(0);
  const [roundUnit, setRoundUnit] = useState(100);
  const [categoryCode, setCategoryCode] = useState("50001770");
  const [courierCode, setCourierCode] = useState("CJGLS");
  const [asPhone, setAsPhone] = useState("01027483227");
  const [multipleOrigins, setMultipleOrigins] = useState<"N" | "Y">("N");
  const [appliedSettings, setAppliedSettings] = useState<AppliedSettings | null>(null);
  const [sourceRows, setSourceRows] = useState<Row[]>([]);
  const [fileName, setFileName] = useState("");
  const [status, setStatus] = useState("상품 일괄목록 엑셀을 업로드해 주세요.");

  const products = useMemo(() => appliedSettings ? toProducts(sourceRows, appliedSettings) : [], [sourceRows, appliedSettings]);

  async function readFile(file?: File) {
    if (!file) return;
    setFileName(file.name);
    setAppliedSettings(null);
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });
      const parsed = parseSheet(workbook.Sheets[workbook.SheetNames[0]]);
      setSourceRows(parsed.rows);
      setStatus(`${parsed.rows.length}행을 읽었습니다. 스마트스토어 탭에서 옵션을 설정하고 적용해 주세요.`);
    } catch {
      setSourceRows([]);
      setStatus("파일을 읽지 못했습니다.");
    }
  }
  function applySettings() {
    if (!sourceRows.length || feeRate < 0 || marginRate < 0) return;
    if (!categoryCode.trim() || !courierCode.trim() || !asPhone.trim()) {
      setAppliedSettings(null);
      setStatus("네이버 카테고리코드·택배사코드·A/S 전화번호를 모두 입력해 주세요.");
      return;
    }
    const next: AppliedSettings = {
      feeRate, marginRate, extraCost, roundUnit,
      categoryCode: categoryCode.trim(), courierCode: courierCode.trim(),
      asPhone: asPhone.trim(), multipleOrigins,
    };
    const preview = toProducts(sourceRows, next);
    const priced = preview.filter((product) => product.basePrice > 0).length;
    if (!priced) {
      setAppliedSettings(null);
      setStatus("가격 열을 찾지 못했습니다.");
      return;
    }
    setAppliedSettings(next);
    setStatus(`스마트스토어 적용 완료: ${priced}개 상품. 이미지1→대표이미지, 이미지2 이후→추가이미지, 상품설명→상세설명으로 반영했습니다.`);
  }
  function downloadNaver() {
    if (!products.length || !appliedSettings) return;
    const dataRows = products.map((product) => productToNaverRow(product, appliedSettings));
    const sheet = XLSX.utils.aoa_to_sheet([groupRow, naverHeaders, ...dataRows]);
    sheet["!cols"] = naverHeaders.map((header) => ({ wch: Math.min(Math.max(header.replace(/\n/g, "").length + 3, 12), 28) }));
    sheet["!rows"] = [{ hpt: 24 }, { hpt: 46 }];
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

  const invalid = feeRate < 0 || marginRate < 0;
  const changed = !!appliedSettings && (
    appliedSettings.feeRate !== feeRate || appliedSettings.marginRate !== marginRate ||
    appliedSettings.extraCost !== extraCost || appliedSettings.roundUnit !== roundUnit ||
    appliedSettings.categoryCode !== categoryCode.trim() || appliedSettings.courierCode !== courierCode.trim() ||
    appliedSettings.asPhone !== asPhone.trim() || appliedSettings.multipleOrigins !== multipleOrigins
  );

  return (
    <main className="container">
      <section className="hero">
        <span className="badge">postsheet02 · 상품 대량등록 변환</span>
        <h1>상품 일괄목록을<br />마켓 등록 파일로 변환</h1>
        <p>엑셀을 한 번 업로드하고 마켓별 탭에서 전용 옵션을 적용해 다운로드합니다.</p>
        <div className="privacy">원본·결과 파일 서버 저장 없음 · 브라우저 안에서만 처리</div>
      </section>

      <section className="panel">
        <div className="field full">
          <label>1. 상품 일괄목록 엑셀 업로드</label>
          <input type="file" accept=".xlsx,.xls,.csv" onChange={(event) => readFile(event.target.files?.[0])} />
          <small>{fileName || "선택된 파일 없음"}</small>
        </div>
      </section>

      <section className="panel">
        <div className="full" style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 16 }}>
          <button type="button" onClick={() => setActiveMarket("smartstore")} style={{ opacity: activeMarket === "smartstore" ? 1 : 0.65 }}>스마트스토어</button>
          <button type="button" disabled style={{ opacity: 0.45 }}>옥션 · 준비 중</button>
          <button type="button" disabled style={{ opacity: 0.45 }}>지마켓 · 준비 중</button>
        </div>

        {activeMarket === "smartstore" && <>
          <div className="field full"><label>스마트스토어 적용 옵션</label><small>아래 값은 스마트스토어 파일에만 적용됩니다.</small></div>
          <div className="field"><label>네이버 수수료율 (%)</label><input min={0} step={0.1} type="number" value={feeRate} onChange={(e) => setFeeRate(Number(e.target.value))} /></div>
          <div className="field"><label>추가 마진율 (%)</label><input min={0} step={0.1} type="number" value={marginRate} onChange={(e) => setMarginRate(Number(e.target.value))} /></div>
          <div className="field"><label>상품당 추가비용 (원)</label><input min={0} type="number" value={extraCost} onChange={(e) => setExtraCost(Number(e.target.value))} /></div>
          <div className="field"><label>판매가 올림 단위</label><select value={roundUnit} onChange={(e) => setRoundUnit(Number(e.target.value))}><option value={10}>10원 올림</option><option value={100}>100원 올림</option><option value={500}>500원 올림</option><option value={1000}>1,000원 올림</option></select></div>
          <div className="field"><label>네이버 카테고리코드</label><input value={categoryCode} onChange={(e) => setCategoryCode(e.target.value)} /></div>
          <div className="field"><label>택배사코드</label><input value={courierCode} onChange={(e) => setCourierCode(e.target.value)} /></div>
          <div className="field"><label>A/S 전화번호</label><input value={asPhone} onChange={(e) => setAsPhone(e.target.value)} /></div>
          <div className="field"><label>복수원산지 여부</label><select value={multipleOrigins} onChange={(e) => setMultipleOrigins(e.target.value as "N" | "Y")}><option value="N">N · 단일 원산지</option><option value="Y">Y · 복수 원산지</option></select></div>
          <div className="field full"><small>이미지1은 대표이미지, 이미지2 이후는 추가이미지, 상품설명은 상세설명으로 자동 연결됩니다.</small></div>
          {changed && <div className="status full">입력값이 변경됐습니다. 다시 ‘위 내용 적용하기’를 눌러 주세요.</div>}
          <div className="actions full">
            <button onClick={downloadNaver} disabled={!products.length || invalid || changed}>스마트스토어 파일 다운로드</button>
            <button onClick={applySettings} disabled={!sourceRows.length || invalid}>위 내용 적용하기</button>
            <button className="secondary" onClick={resetAll}>초기화</button>
          </div>
          <div className="status full">{status}</div>
        </>}
      </section>

      <section className="preview">
        <div className="previewHead"><h2>스마트스토어 가격 계산 미리보기</h2><span>{products.length}개 상품</span></div>
        <div className="tableWrap"><table><thead><tr><th>상품명</th><th>상품코드</th><th>원래 판매가격</th><th>최종 판매가</th><th>대표이미지</th><th>추가이미지</th></tr></thead><tbody>
          {products.slice(0, 30).map((product, index) => <tr key={`${product.sellerCode}-${index}`}><td>{product.productName}</td><td>{product.sellerCode}</td><td>{product.basePrice.toLocaleString()}</td><td>{product.salePrice.toLocaleString()}</td><td>{product.mainImage ? "있음" : "없음"}</td><td>{product.additionalImages ? "있음" : "없음"}</td></tr>)}
          {!products.length && <tr><td colSpan={6} className="empty">엑셀 업로드 후 스마트스토어 탭에서 옵션을 입력하고 적용해 주세요.</td></tr>}
        </tbody></table></div>
      </section>
    </main>
  );
}

"use client";

import { useMemo, useState } from "react";
import * as XLSX from "xlsx";
import JSZip from "jszip";

type Row = Record<string, unknown>;
type MarketTab = "smartstore" | "esm";
type CourierName = "CJ대한통운" | "한진택배";

type Product = {
  raw: Row;
  productName: string;
  sellerCode: string;
  basePrice: number;
  stock: number;
  optionName: string;
  optionValue: string;
  mainImage: string;
  additionalImage: string;
  detailHtml: string;
  shippingFee: number;
  vatType: string;
  originCode: string;
  originDirect: string;
};

type SmartSettings = {
  feeRate: number;
  marginRate: number;
  extraCost: number;
  roundUnit: number;
  categoryCode: string;
  courierCode: string;
  asPhone: string;
  multipleOrigins: "N" | "Y";
};

type EsmSettings = {
  marginRate: number;
  roundUnit: number;
  auctionId: string;
  gmarketId: string;
  courierName: CourierName;
};

const aliases = {
  productName: ["상품명", "제품명", "상품이름", "품명", "상품 명"],
  sellerCode: ["판매자 상품코드", "판매자상품코드", "상품코드", "자체상품코드", "관리코드", "품목코드", "판매자코드"],
  basePrice: ["판매가", "판매가격", "상품가격", "상품 판매가", "공급가", "공급가격", "판매단가", "단가", "기준가격", "원가", "매입가", "소비자가", "정상가", "가격"],
  stock: ["재고수량", "재고", "수량", "판매가능수량", "A 재고", "G 재고"],
  optionName: ["옵션명", "옵션항목", "옵션"],
  optionValue: ["옵션값", "옵션내용", "선택옵션", "옵션 입력값"],
  mainImage: ["대표이미지", "대표이미지URL", "이미지1", "메인이미지", "이미지URL", "기본이미지"],
  detailHtml: ["상세설명", "상품설명", "상세HTML", "상품상세", "상세페이지", "상세이미지", "상품상세설명"],
  shippingFee: ["기본배송비", "배송비", "반품/교환 배송비"],
  vatType: ["부가세", "과세구분", "부가세유형", "부가세여부"],
  originCode: ["원산지코드", "원산지 지역코드"],
  originDirect: ["원산지 직접입력", "원산지직접입력", "원산지", "원산지명"],
};

const naverHeaders = [
  "판매자 상품코드","카테고리코드","상품명","상품상태","판매가","단위가격 사용여부","표시용량","표시단위","총용량","부가세","관부가세","재고수량","옵션형태","옵션명","옵션값","옵션가","옵션 재고수량","직접입력 옵션","추가상품명","추가상품값","추가상품가","추가상품 재고수량","대표이미지","추가이미지","상세설명","브랜드","제조사","제조일자","유효일자","원산지코드","수입사","복수원산지여부","원산지 직접입력","미성년자 구매","배송비 템플릿코드","배송방법","택배사코드","배송비유형","기본배송비","배송비 결제방식","조건부무료-\n상품판매가 합계","수량별부과-수량","구간별-\n2구간수량","구간별-\n3구간수량","구간별-\n3구간배송비","구간별-\n추가배송비","반품배송비","교환배송비","지역별 차등 배송비","별도설치비","상품정보제공고시 템플릿코드","상품정보제공고시\n품명","상품정보제공고시\n모델명","상품정보제공고시\n인증허가사항","상품정보제공고시\n제조자","A/S 템플릿코드","A/S 전화번호","A/S 안내","판매자특이사항","즉시할인 값\n(기본할인)","즉시할인 단위\n(기본할인)","모바일\n즉시할인 값","모바일\n즉시할인 단위","복수구매할인\n조건 값","복수구매할인\n조건 단위","복수구매할인\n값","복수구매할인\n단위","상품구매시 포인트\n지급 값","상품구매시 포인트\n지급 단위","텍스트리뷰 작성시\n지급 포인트","포토/동영상 리뷰 작성시\n지급 포인트","한달사용 텍스트리뷰\n작성시 지급 포인트","한달사용\n포토/동영상리뷰 작성시 지급 포인트","알림받기동의 고객 리뷰 작성 시 지급 포인트","무이자\n할부 개월","사은품","판매자바코드","구매평 노출여부","구매평\n비노출사유","알림받기 동의 고객 전용 여부","ISBN","ISSN","독립출판","출간일","출판사","글작가","그림작가","번역자명","문화비 소득공제","사이즈\n상품군","사이즈\n사이즈명","사이즈\n상세 사이즈","사이즈\n모델명","판매준수가"
];

const naverGroupRow = naverHeaders.map((_, i) =>
  i === 0 ? "상품 기본정보" :
  i === 25 ? "상품 주요정보" :
  i === 34 ? "배송정보" :
  i === 50 ? "상품정보제공고시" :
  i === 55 ? "A/S, 특이사항" :
  i === 59 ? "할인/혜택정보" :
  i === 77 ? "기타 정보" : ""
);

const esmHeaders = ["노출\n사이트","A ID","G ID","상품명","A프로모션 문구","G프로모션 문구","G 영문","G 중문","카테고리 템플릿 코드","카테고리 코드","A 노출코드","G 노출코드","판매기간","A 판매가","G 판매가","A 할인유형","A 할인가","G 할인유형","G 할인가","A 재고","G 재고","옵션\n타입","옵션명","옵션\n입력값","기본이미지","추가이미지","상품상세설명","배송정보 \n템플릿 코드","배송방법","출하지 코드","배송정책번호","반품/교환\n주소 코드","A 발송정책","G 발송정책","택배사\n코드","반품/교환\n배송비","상품군\n코드","상품고시정보\n템플릿코드","인증타입","인증품목선택","인증코드","인증타입","인증품목선택","인증코드","병행수입여부","인증타입","인증품목선택","인증코드","병행수입여부","인증타입","승인/신고번호","원산지\n상품타입","원산지\n지역타입","원산지\n지역코드","복수\n원산지여부","사은품/덤 \n템플릿 코드","사은품","덤","소비기한","제조일자","청소년구매\n불가여부","부가세여부","선물하기상품"];

const esmTopRows: unknown[][] = [
  ["▶가이드 바로가기", null, " ※ 문서버전 : NEW 2.0", null, "필독▶ 상품정보는 8행부터 입력합니다. 1~7행은 삭제하지 마세요."],
  [null, "상품기본정보", ...Array(26).fill(null), "배송정보", ...Array(8).fill(null), "상품고시정보", ...Array(25).fill(null), "추가정보"],
  [null, "계정선택", null, null, "상품명", null, null, null, null, "카테고리", null, null, null, "판매가", null, "할인", null, null, null, "재고수량", null, " 옵션", null, null, "상품이미지", null, null, null, null, "배송정보"],
  [null, ...esmHeaders],
  [null, ...esmHeaders.map((_, i) => [0,1,2,3,9,10,11,12,13,14,19,20,24,26,28,29,30,31,32,33,34,35,36,37,51,52,53,54,60,61].includes(i) ? "필수" : "비필수")],
  [null, ...esmHeaders.map(() => "")],
  ["예시", "옥션/G마켓", "auctionid", "gmarketid", "상품명", "", "", "", "", "", "", "", "", "무제한", 100000, 100000, "", "", "", "", 99999, 99999, "미사용", "", "", "https://example.com/1.jpg", "https://example.com/2.jpg", "<p>상품설명</p>", "", "일반택배", "", "", "", "", "", 10013, 2500, "", "", "인증대상아님", "", "", "인증대상아님", "", "", "해당사항없음", "인증대상아님", "", "", "해당사항없음", "인증대상아님", "", "해당없음", "알수없음", "", "단일원산지", "", "", "", "", "", "구매가능", "과세상품", "가능"]
];

function cleanKey(value: unknown) {
  return String(value ?? "").replace(/[\s\n\r_\-()\[\]\/]/g, "").toLowerCase();
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
    const exact = pickExact(row, name);
    if (normalize(exact) !== "") return exact;
  }
  for (const name of names) {
    const target = cleanKey(name);
    const found = Object.entries(row).find(([key]) => {
      const cleaned = cleanKey(key);
      return cleaned && (cleaned.includes(target) || target.includes(cleaned));
    });
    if (found && normalize(found[1]) !== "") return found[1];
  }
  return "";
}
function getAdditionalImage(row: Row) {
  return normalize(pickExact(row, "이미지2")) || normalize(pickExact(row, "추가이미지"));
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
  return matrix.slice(bestIndex + 1)
    .map((values) => Object.fromEntries(headers.map((header, index) => [header, values[index] ?? ""])))
    .filter((row) => Object.values(row).some((value) => normalize(value) !== ""));
}
function toProducts(rows: Row[]): Product[] {
  return rows.map((raw) => ({
    raw,
    productName: normalize(pick(raw, aliases.productName)),
    sellerCode: normalize(pick(raw, aliases.sellerCode)),
    basePrice: numberValue(pick(raw, aliases.basePrice)),
    stock: numberValue(pick(raw, aliases.stock)) || 99999,
    optionName: normalize(pick(raw, aliases.optionName)),
    optionValue: normalize(pick(raw, aliases.optionValue)),
    mainImage: normalize(pick(raw, aliases.mainImage)),
    additionalImage: getAdditionalImage(raw),
    detailHtml: normalize(pick(raw, aliases.detailHtml)),
    shippingFee: numberValue(pick(raw, aliases.shippingFee)),
    vatType: normalize(pick(raw, aliases.vatType)),
    originCode: normalize(pick(raw, aliases.originCode)),
    originDirect: normalize(pick(raw, aliases.originDirect)),
  })).filter((product) => (product.productName || product.sellerCode) && product.basePrice > 0);
}
function roundedPrice(base: number, marginRate: number, roundUnit: number, extraCost = 0) {
  if (base <= 0) return 0;
  const calculated = (base + extraCost) * (1 + marginRate / 100);
  return Math.ceil(calculated / Math.max(roundUnit, 1)) * Math.max(roundUnit, 1);
}
function copyCode(raw: Row, names: string[]) { return normalize(pick(raw, names)); }
function courierCode(name: CourierName) { return name === "CJ대한통운" ? 10013 : 10007; }
function esmOptionType(product: Product) {
  if (!product.optionName || !product.optionValue) return "미사용";
  const count = product.optionName.split(/[,\n]/).filter(Boolean).length;
  return count >= 3 ? "3개조합형" : count === 2 ? "2개조합형" : "단독형";
}
function downloadBlob(blob: Blob, name: string) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = name;
  anchor.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function esmRow(product: Product, settings: EsmSettings) {
  const row = Array(esmHeaders.length).fill("");
  const set = (header: string, value: unknown) => {
    const index = esmHeaders.indexOf(header);
    if (index >= 0) row[index] = value ?? "";
  };
  const raw = product.raw;
  const finalPrice = roundedPrice(product.basePrice, settings.marginRate, settings.roundUnit);

  set("노출\n사이트", "옥션/G마켓");
  set("A ID", settings.auctionId);
  set("G ID", settings.gmarketId);
  set("상품명", product.productName);
  set("카테고리 템플릿 코드", copyCode(raw, ["카테고리 템플릿 코드", "카테고리템플릿코드"]));
  set("카테고리 코드", copyCode(raw, ["ESM 카테고리코드", "카테고리 코드", "카테고리코드"]));
  set("A 노출코드", copyCode(raw, ["A 노출코드", "옥션 노출코드"]));
  set("G 노출코드", copyCode(raw, ["G 노출코드", "G마켓 노출코드"]));
  set("판매기간", "무제한");
  set("A 판매가", finalPrice);
  set("G 판매가", finalPrice);
  set("A 재고", product.stock);
  set("G 재고", product.stock);
  set("옵션\n타입", esmOptionType(product));
  set("옵션명", product.optionName);
  set("옵션\n입력값", product.optionValue);
  set("기본이미지", product.mainImage);
  set("추가이미지", product.additionalImage);
  set("상품상세설명", product.detailHtml);
  set("배송정보 \n템플릿 코드", copyCode(raw, ["배송정보 템플릿 코드", "배송정보템플릿코드"]));
  set("배송방법", "일반택배");
  set("출하지 코드", copyCode(raw, ["출하지 코드", "출하지코드"]));
  set("배송정책번호", copyCode(raw, ["배송정책번호"]));
  set("반품/교환\n주소 코드", copyCode(raw, ["반품/교환 주소 코드", "반품교환주소코드"]));
  set("A 발송정책", copyCode(raw, ["A 발송정책", "옥션 발송정책"]));
  set("G 발송정책", copyCode(raw, ["G 발송정책", "G마켓 발송정책"]));
  set("택배사\n코드", courierCode(settings.courierName));
  set("반품/교환\n배송비", numberValue(pick(raw, ["반품/교환 배송비", "반품배송비", "교환배송비"])) || product.shippingFee || 2500);
  set("상품군\n코드", copyCode(raw, ["상품군 코드", "상품군코드"]));
  set("상품고시정보\n템플릿코드", copyCode(raw, ["상품고시정보 템플릿코드", "상품고시 템플릿코드"]));
  set("인증타입", "인증대상아님");
  set("병행수입여부", "해당사항없음");
  set("원산지\n상품타입", copyCode(raw, ["원산지 상품타입"]) || "해당없음");
  set("원산지\n지역타입", copyCode(raw, ["원산지 지역타입"]) || "알수없음");
  set("원산지\n지역코드", product.originCode);
  set("복수\n원산지여부", copyCode(raw, ["복수원산지여부"]) || "단일원산지");
  set("청소년구매\n불가여부", "구매가능");
  set("부가세여부", product.vatType.includes("면세") ? "면세상품" : "과세상품");
  set("선물하기상품", "가능");
  return ["", ...row];
}
function makeEsmWorkbook(products: Product[], settings: EsmSettings) {
  const data = [...esmTopRows, ...products.map((product) => esmRow(product, settings))];
  const sheet = XLSX.utils.aoa_to_sheet(data);
  sheet["!cols"] = Array(64).fill({ wch: 16 });
  sheet["!rows"] = [{hpt:24},{hpt:24},{hpt:24},{hpt:50},{hpt:24},{hpt:120},{hpt:32}];
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet, "NEW 일반상품");
  return workbook;
}

export default function Home() {
  const [active, setActive] = useState<MarketTab>("smartstore");
  const [sourceRows, setSourceRows] = useState<Row[]>([]);
  const [fileName, setFileName] = useState("");
  const [status, setStatus] = useState("상품 엑셀을 업로드해 주세요.");
  const products = useMemo(() => toProducts(sourceRows), [sourceRows]);

  const [feeRate, setFeeRate] = useState(6);
  const [marginRate, setMarginRate] = useState(30);
  const [extraCost, setExtraCost] = useState(0);
  const [roundUnit, setRoundUnit] = useState(100);
  const [categoryCode, setCategoryCode] = useState("50001770");
  const [naverCourier, setNaverCourier] = useState("CJGLS");
  const [asPhone, setAsPhone] = useState("01027483227");
  const [multipleOrigins, setMultipleOrigins] = useState<"N" | "Y">("N");
  const [smartApplied, setSmartApplied] = useState<SmartSettings | null>(null);

  const [esmMargin, setEsmMargin] = useState(30);
  const [esmRound, setEsmRound] = useState(100);
  const [auctionId, setAuctionId] = useState("");
  const [gmarketId, setGmarketId] = useState("");
  const [esmCourier, setEsmCourier] = useState<CourierName>("CJ대한통운");
  const [esmApplied, setEsmApplied] = useState<EsmSettings | null>(null);

  const esmPreview = useMemo(() => products.slice(0, 30).map((product) => ({
    ...product,
    finalPrice: roundedPrice(product.basePrice, esmMargin, esmRound),
  })), [products, esmMargin, esmRound]);

  async function readFile(file?: File) {
    if (!file) return;
    setFileName(file.name);
    setSmartApplied(null);
    setEsmApplied(null);
    try {
      const workbook = XLSX.read(await file.arrayBuffer(), { type: "array" });
      const rows = parseSheet(workbook.Sheets[workbook.SheetNames[0]]);
      const parsedProducts = toProducts(rows);
      setSourceRows(rows);
      const priced = parsedProducts.filter((product) => product.basePrice > 0).length;
      setStatus(`${rows.length}행을 읽었고 판매가가 확인된 상품은 ${priced}개입니다.`);
    } catch {
      setSourceRows([]);
      setStatus("파일을 읽지 못했습니다.");
    }
  }

  function applySmart() {
    if (!products.length) return;
    setSmartApplied({ feeRate, marginRate, extraCost, roundUnit, categoryCode: categoryCode.trim(), courierCode: naverCourier.trim(), asPhone: asPhone.trim(), multipleOrigins });
    setStatus(`스마트스토어 ${products.length}개 적용 완료`);
  }
  function smartRow(product: Product, settings: SmartSettings) {
    const row = naverHeaders.map((header) => pickExact(product.raw, header));
    const set = (header: string, value: unknown) => {
      const index = naverHeaders.indexOf(header);
      if (index >= 0) row[index] = value ?? "";
    };
    set("판매자 상품코드", product.sellerCode);
    set("카테고리코드", settings.categoryCode);
    set("상품명", product.productName);
    set("상품상태", "신상품");
    set("판매가", roundedPrice(product.basePrice, settings.marginRate + settings.feeRate, settings.roundUnit, settings.extraCost));
    set("단위가격 사용여부", "N");
    set("부가세", product.vatType || "과세상품");
    set("재고수량", product.stock === 99999 ? 1 : product.stock);
    set("대표이미지", product.mainImage);
    set("추가이미지", product.additionalImage);
    set("상세설명", product.detailHtml);
    set("원산지코드", "03");
    set("복수원산지여부", settings.multipleOrigins);
    set("원산지 직접입력", product.originDirect);
    set("미성년자 구매", "Y");
    set("배송방법", "택배, 소포, 등기");
    set("택배사코드", settings.courierCode);
    set("배송비유형", product.shippingFee > 0 ? "유료" : "무료");
    set("기본배송비", product.shippingFee);
    set("반품배송비", product.shippingFee || 3000);
    set("교환배송비", (product.shippingFee || 3000) * 2);
    set("별도설치비", "N");
    set("A/S 전화번호", settings.asPhone);
    set("A/S 안내", "판매자에 문의하시거나, A/S연락처로 문의 주시기 바랍니다.");
    set("구매평 노출여부", "Y");
    set("알림받기 동의 고객 전용 여부", "N");
    return row;
  }
  function downloadSmart() {
    if (!smartApplied) return;
    const sheet = XLSX.utils.aoa_to_sheet([naverGroupRow, naverHeaders, ...products.map((product) => smartRow(product, smartApplied))]);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, sheet, "일괄등록");
    XLSX.writeFile(workbook, `postsheet02_스마트스토어_${new Date().toISOString().slice(0, 10)}.xlsx`);
  }

  function applyEsm() {
    if (!products.length) return;
    if (!auctionId.trim() && !gmarketId.trim()) {
      setStatus("옥션 또는 G마켓 판매자 ID를 하나 이상 입력해 주세요.");
      return;
    }
    const settings: EsmSettings = { marginRate: esmMargin, roundUnit: esmRound, auctionId: auctionId.trim(), gmarketId: gmarketId.trim(), courierName: esmCourier };
    setEsmApplied(settings);
    const sample = products[0];
    const final = sample ? roundedPrice(sample.basePrice, esmMargin, esmRound) : 0;
    setStatus(`ESM 적용 완료 · 예시 ${sample?.basePrice.toLocaleString() ?? 0}원 → ${final.toLocaleString()}원 · ${Math.ceil(products.length / 500)}개 파일`);
  }
  async function downloadEsm() {
    if (!esmApplied) return;
    const chunks: Product[][] = [];
    for (let index = 0; index < products.length; index += 500) chunks.push(products.slice(index, index + 500));
    const date = new Date().toISOString().slice(0, 10);
    if (chunks.length === 1) {
      XLSX.writeFile(makeEsmWorkbook(chunks[0], esmApplied), `ESM_${chunks[0].length}개_${date}.xlsx`);
      return;
    }
    const zip = new JSZip();
    chunks.forEach((chunk, index) => {
      const output = XLSX.write(makeEsmWorkbook(chunk, esmApplied), { bookType: "xlsx", type: "array" });
      zip.file(`ESM_${String(index + 1).padStart(2, "0")}_${chunk.length}개.xlsx`, output);
    });
    downloadBlob(await zip.generateAsync({ type: "blob" }), `ESM_${products.length}개_분할파일_${date}.zip`);
  }

  return <main className="container">
    <section className="hero">
      <span className="badge">postsheet02 · 상품 대량등록 변환</span>
      <h1>상품 일괄목록을<br />마켓 등록 파일로 변환</h1>
      <p>업로드 파일을 자동 해석해 스마트스토어와 ESM 양식으로 변환합니다.</p>
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
        <button onClick={() => setActive("smartstore")} style={{ opacity: active === "smartstore" ? 1 : 0.6 }}>스마트스토어</button>
        <button onClick={() => setActive("esm")} style={{ opacity: active === "esm" ? 1 : 0.6 }}>ESM · 옥션/G마켓</button>
      </div>

      {active === "smartstore" && <>
        <div className="field"><label>네이버 수수료율 (%)</label><input type="number" value={feeRate} onChange={(event) => setFeeRate(Number(event.target.value))} /></div>
        <div className="field"><label>추가 마진율 (%)</label><input type="number" value={marginRate} onChange={(event) => setMarginRate(Number(event.target.value))} /></div>
        <div className="field"><label>상품당 추가비용</label><input type="number" value={extraCost} onChange={(event) => setExtraCost(Number(event.target.value))} /></div>
        <div className="field"><label>판매가 올림 단위</label><select value={roundUnit} onChange={(event) => setRoundUnit(Number(event.target.value))}><option value={1}>1원</option><option value={10}>10원</option><option value={100}>100원</option><option value={500}>500원</option><option value={1000}>1,000원</option></select></div>
        <div className="field"><label>카테고리코드</label><input value={categoryCode} onChange={(event) => setCategoryCode(event.target.value)} /></div>
        <div className="field"><label>택배사코드</label><input value={naverCourier} onChange={(event) => setNaverCourier(event.target.value)} /></div>
        <div className="field"><label>A/S 전화번호</label><input value={asPhone} onChange={(event) => setAsPhone(event.target.value)} /></div>
        <div className="field"><label>복수원산지</label><select value={multipleOrigins} onChange={(event) => setMultipleOrigins(event.target.value as "N" | "Y")}><option value="N">단일</option><option value="Y">복수</option></select></div>
        <div className="actions full"><button onClick={downloadSmart} disabled={!smartApplied}>스마트스토어 다운로드</button><button onClick={applySmart} disabled={!products.length}>위 내용 적용하기</button></div>
      </>}

      {active === "esm" && <>
        <div className="field full"><label>ESM 자동 변환</label><small>마진율은 원본 판매가에 즉시 적용되며 아래 미리보기에서 계산 결과를 확인할 수 있습니다.</small></div>
        <div className="field"><label>마진율 (%)</label><input type="number" min={0} value={esmMargin} onChange={(event) => { setEsmMargin(Number(event.target.value)); setEsmApplied(null); }} /></div>
        <div className="field"><label>판매가 올림 단위</label><select value={esmRound} onChange={(event) => { setEsmRound(Number(event.target.value)); setEsmApplied(null); }}><option value={1}>1원</option><option value={10}>10원</option><option value={100}>100원</option><option value={500}>500원</option><option value={1000}>1,000원</option></select></div>
        <div className="field"><label>옥션 판매자 ID</label><input value={auctionId} onChange={(event) => { setAuctionId(event.target.value); setEsmApplied(null); }} /></div>
        <div className="field"><label>G마켓 판매자 ID</label><input value={gmarketId} onChange={(event) => { setGmarketId(event.target.value); setEsmApplied(null); }} /></div>
        <div className="field"><label>발송 택배사</label><select value={esmCourier} onChange={(event) => { setEsmCourier(event.target.value as CourierName); setEsmApplied(null); }}><option value="CJ대한통운">CJ대한통운 · 10013</option><option value="한진택배">한진택배 · 10007</option></select></div>
        <div className="field full"><small>{products.length > 500 ? `${products.length}개 상품을 500개 단위로 ${Math.ceil(products.length / 500)}개 엑셀에 나누어 ZIP으로 다운로드합니다.` : "500개 이하이므로 엑셀 한 개로 다운로드합니다."}</small></div>
        <div className="actions full"><button onClick={downloadEsm} disabled={!esmApplied}>ESM 파일 다운로드</button><button onClick={applyEsm} disabled={!products.length}>위 내용 적용하기</button></div>
      </>}
      <div className="status full">{status}</div>
    </section>

    <section className="preview">
      <div className="previewHead"><h2>{active === "esm" ? "ESM 가격 계산 미리보기" : "자동 해석 미리보기"}</h2><span>{products.length}개 상품</span></div>
      <div className="tableWrap"><table><thead><tr><th>상품명</th><th>원래 판매가</th>{active === "esm" && <th>마진 적용 판매가</th>}<th>재고</th><th>이미지1</th><th>이미지2</th></tr></thead><tbody>
        {(active === "esm" ? esmPreview : products.slice(0, 30)).map((product, index) => <tr key={`${product.sellerCode}-${index}`}><td>{product.productName}</td><td>{product.basePrice.toLocaleString()}</td>{active === "esm" && <td>{("finalPrice" in product ? product.finalPrice : 0).toLocaleString()}</td>}<td>{product.stock}</td><td>{product.mainImage ? "있음" : "없음"}</td><td>{product.additionalImage ? "있음" : "없음"}</td></tr>)}
        {!products.length && <tr><td colSpan={active === "esm" ? 6 : 5} className="empty">파일을 업로드해 주세요.</td></tr>}
      </tbody></table></div>
    </section>
  </main>;
}

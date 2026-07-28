"use client";

import { useMemo, useState } from "react";
import * as XLSX from "xlsx";
import JSZip from "jszip";

type Row = Record<string, unknown>;
type MarketTab = "smartstore" | "esm";
type StepId = "category" | "notice" | "origin" | "shipping" | "price" | "review";
type CourierName = "CJ대한통운" | "한진택배";

type Product = {
  id: string;
  raw: Row;
  productName: string;
  sellerCode: string;
  basePrice: number;
  finalPrice: number;
  stock: number;
  optionName: string;
  optionValue: string;
  mainImage: string;
  additionalImage: string;
  detailHtml: string;
  shippingFee: number;
  vatType: string;

  categoryGroup: string;
  categoryCode: string;
  auctionExposureCode: string;
  gmarketExposureCode: string;
  productGroupCode: string;

  noticeTemplateCode: string;

  originGroup: string;
  originProductType: string;
  originRegionType: string;
  originRegionCode: string;
  multipleOrigins: string;

  shippingGroup: string;
  departureCode: string;
  shippingPolicyNumber: string;
  returnAddressCode: string;
  auctionShippingPolicy: string;
  gmarketShippingPolicy: string;
  courierName: CourierName;
  courierCode: string;
  returnShippingFee: number;
};

type ValidationIssue = {
  productId: string;
  productName: string;
  field: string;
  message: string;
  step: StepId;
};

type NoticeRule = { code: string; name: string; keywords: string[] };

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

const noticeRules: NoticeRule[] = [
  { code: "1", name: "의류", keywords: ["티셔츠", "셔츠", "블라우스", "바지", "스커트", "원피스", "자켓", "재킷", "코트", "의류", "속옷", "양말", "잠옷"] },
  { code: "2", name: "구두/신발", keywords: ["신발", "구두", "운동화", "스니커즈", "슬리퍼", "샌들", "부츠", "로퍼"] },
  { code: "3", name: "가방", keywords: ["가방", "백팩", "크로스백", "토트백", "파우치", "캐리어"] },
  { code: "17", name: "주방용품", keywords: ["냄비", "프라이팬", "칼", "도마", "수저", "그릇", "컵", "텀블러", "주방", "밀폐용기", "조리도구"] },
  { code: "18", name: "화장품", keywords: ["화장품", "크림", "로션", "에센스", "세럼", "립스틱", "샴푸", "린스", "클렌징", "선크림"] },
  { code: "20", name: "농수축산물", keywords: ["사과", "배", "과일", "채소", "농산물", "수산물", "생선", "고기", "축산물", "쌀", "계란"] },
  { code: "21", name: "가공식품", keywords: ["과자", "라면", "커피", "차", "음료", "소스", "통조림", "가공식품", "만두", "떡", "빵", "초콜릿"] },
  { code: "22", name: "건강기능식품", keywords: ["건강기능식품", "비타민", "유산균", "오메가", "홍삼", "영양제"] },
  { code: "35", name: "기타 재화", keywords: [] },
];

const naverHeaders = ["판매자 상품코드", "카테고리코드", "상품명", "상품상태", "판매가", "부가세", "재고수량", "대표이미지", "추가이미지", "상세설명", "원산지코드", "복수원산지여부", "원산지 직접입력", "미성년자 구매", "배송방법", "택배사코드", "배송비유형", "기본배송비", "반품배송비", "교환배송비", "별도설치비", "A/S 전화번호", "A/S 안내", "구매평 노출여부", "알림받기 동의 고객 전용 여부"];

const esmHeaders = ["노출\n사이트", "A ID", "G ID", "상품명", "A프로모션 문구", "G프로모션 문구", "G 영문", "G 중문", "카테고리 템플릿 코드", "카테고리 코드", "A 노출코드", "G 노출코드", "판매기간", "A 판매가", "G 판매가", "A 할인유형", "A 할인가", "G 할인유형", "G 할인가", "A 재고", "G 재고", "옵션\n타입", "옵션명", "옵션\n입력값", "기본이미지", "추가이미지", "상품상세설명", "배송정보 \n템플릿 코드", "배송방법", "출하지 코드", "배송정책번호", "반품/교환\n주소 코드", "A 발송정책", "G 발송정책", "택배사\n코드", "반품/교환\n배송비", "상품군\n코드", "상품고시정보\n템플릿코드", "인증타입", "인증품목선택", "인증코드", "인증타입", "인증품목선택", "인증코드", "병행수입여부", "인증타입", "인증품목선택", "인증코드", "병행수입여부", "인증타입", "승인/신고번호", "원산지\n상품타입", "원산지\n지역타입", "원산지\n지역코드", "복수\n원산지여부", "사은품/덤 \n템플릿 코드", "사은품", "덤", "소비기한", "제조일자", "청소년구매\n불가여부", "부가세여부", "선물하기상품"];

const steps: { id: StepId; label: string }[] = [
  { id: "category", label: "1. 카테고리" },
  { id: "notice", label: "2. 고시정보" },
  { id: "origin", label: "3. 원산지" },
  { id: "shipping", label: "4. 배송정책" },
  { id: "price", label: "5. 가격·배송비" },
  { id: "review", label: "6. 최종검사" },
];

function cleanKey(v: unknown) {
  return String(v ?? "").replace(/[\s\n\r_\-()\[\]\/]/g, "").toLowerCase();
}
function normalize(v: unknown) { return String(v ?? "").trim(); }
function digits(v: unknown) { return normalize(v).replace(/\D/g, ""); }
function numberValue(v: unknown) {
  const n = Number(String(v ?? "").replace(/,/g, "").replace(/[^0-9.-]/g, ""));
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
  return "";
}
function parseSheet(sheet: XLSX.WorkSheet) {
  const matrix = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "", raw: false });
  const candidates = [...aliases.productName, ...aliases.basePrice].map(cleanKey);
  let bestIndex = 0;
  let bestScore = -1;
  matrix.slice(0, 30).forEach((row, index) => {
    const score = row.map(cleanKey).filter(Boolean).reduce((sum, key) => sum + (candidates.some(c => key === c || key.includes(c) || c.includes(key)) ? 1 : 0), 0);
    if (score > bestScore) { bestScore = score; bestIndex = index; }
  });
  const headers = (matrix[bestIndex] ?? []).map((value, index) => normalize(value) || `열${index + 1}`);
  return matrix.slice(bestIndex + 1)
    .map(values => Object.fromEntries(headers.map((header, index) => [header, values[index] ?? ""])))
    .filter(row => Object.values(row).some(value => normalize(value) !== ""));
}
function detectCategory(row: Row, productName: string) {
  const existing = normalize(pickExact(row, "카테고리코드"));
  if (existing) return { group: `기존 카테고리 ${existing}`, code: existing, productGroupCode: normalize(pickExact(row, "상품군 코드")) };
  const text = `${productName} ${Object.values(row).join(" ")}`.toLowerCase();
  const rule = noticeRules.find(item => item.keywords.some(keyword => text.includes(keyword.toLowerCase()))) ?? noticeRules[noticeRules.length - 1];
  return { group: `${rule.name} 예상`, code: "", productGroupCode: rule.code };
}
function detectOrigin(row: Row) {
  const direct = normalize(pick(row, aliases.originDirect));
  const code = normalize(pick(row, aliases.originCode));
  const group = direct || code || "원산지 미확인";
  return { group, code };
}
function courierCode(name: CourierName) { return name === "한진택배" ? "10007" : "10013"; }
function roundedPrice(base: number, margin: number, unit: number) {
  if (margin === 0) return base;
  return Math.ceil((base * (1 + margin / 100)) / Math.max(1, unit)) * Math.max(1, unit);
}
function makeProducts(rows: Row[]): Product[] {
  return rows.map((raw, index) => {
    const productName = normalize(pick(raw, aliases.productName));
    const basePrice = numberValue(pick(raw, aliases.basePrice));
    const category = detectCategory(raw, productName);
    const origin = detectOrigin(raw);
    const courierName: CourierName = normalize(pickExact(raw, "택배사")) === "한진택배" ? "한진택배" : "CJ대한통운";
    const shippingKey = [normalize(pickExact(raw, "출하지 코드")), normalize(pickExact(raw, "배송정책번호")), normalize(pickExact(raw, "A 발송정책")), normalize(pickExact(raw, "G 발송정책"))].filter(Boolean).join(" / ") || "배송정보 미입력";
    return {
      id: `P${index + 1}`,
      raw,
      productName,
      sellerCode: normalize(pick(raw, aliases.sellerCode)) || `P${index + 1}`,
      basePrice,
      finalPrice: basePrice,
      stock: numberValue(pick(raw, aliases.stock)) || 99999,
      optionName: normalize(pick(raw, aliases.optionName)),
      optionValue: normalize(pick(raw, aliases.optionValue)),
      mainImage: normalize(pick(raw, aliases.mainImage)),
      additionalImage: normalize(pickExact(raw, "이미지2")) || normalize(pickExact(raw, "추가이미지")),
      detailHtml: normalize(pick(raw, aliases.detailHtml)),
      shippingFee: numberValue(pick(raw, aliases.shippingFee)),
      vatType: normalize(pick(raw, aliases.vatType)),

      categoryGroup: category.group,
      categoryCode: category.code,
      auctionExposureCode: normalize(pickExact(raw, "A 노출코드")),
      gmarketExposureCode: normalize(pickExact(raw, "G 노출코드")),
      productGroupCode: category.productGroupCode,

      noticeTemplateCode: normalize(pickExact(raw, "상품고시정보 템플릿코드")) || "239479",

      originGroup: origin.group,
      originProductType: normalize(pickExact(raw, "원산지 상품타입")) || (origin.code ? "해당없음" : "해당없음"),
      originRegionType: normalize(pickExact(raw, "원산지 지역타입")) || (origin.code ? "국내산" : "알수없음"),
      originRegionCode: origin.code,
      multipleOrigins: normalize(pickExact(raw, "복수 원산지여부")) || "단일원산지",

      shippingGroup: shippingKey,
      departureCode: normalize(pickExact(raw, "출하지 코드")),
      shippingPolicyNumber: normalize(pickExact(raw, "배송정책번호")),
      returnAddressCode: normalize(pickExact(raw, "반품/교환 주소 코드")),
      auctionShippingPolicy: normalize(pickExact(raw, "A 발송정책")),
      gmarketShippingPolicy: normalize(pickExact(raw, "G 발송정책")),
      courierName,
      courierCode: normalize(pickExact(raw, "택배사 코드")) || courierCode(courierName),
      returnShippingFee: numberValue(pickExact(raw, "반품/교환 배송비")) || numberValue(pick(raw, aliases.shippingFee)) || 2500,
    };
  }).filter(product => product.productName && product.basePrice > 0);
}

function groupProducts(products: Product[], step: StepId) {
  const keyFor = (product: Product) => {
    if (step === "category") return product.categoryGroup;
    if (step === "notice") return `${product.categoryCode || "카테고리 미입력"} / ${product.noticeTemplateCode || "고시코드 미입력"}`;
    if (step === "origin") return product.originGroup;
    if (step === "shipping") return product.shippingGroup;
    return "전체 상품";
  };
  const map = new Map<string, Product[]>();
  products.forEach(product => {
    const key = keyFor(product);
    map.set(key, [...(map.get(key) ?? []), product]);
  });
  return [...map.entries()].map(([name, items]) => ({ name, items }));
}

function esmTopRows(): unknown[][] {
  return [
    ["▶가이드 바로가기", null, " ※ 문서버전 : NEW 2.0", null, "필독▶ 일반배송 전용 파일입니다. 상품정보는 8행부터 입력됩니다.", ...Array(59).fill(null)],
    [null, "상품기본정보", ...Array(26).fill(null), "배송정보", ...Array(8).fill(null), "상품고시정보", "공통 안전인증정보", ...Array(12).fill(null), "추가정보", ...Array(11).fill(null)],
    [null, "계정선택", null, null, "상품명", null, null, null, null, "카테고리", null, null, null, null, "판매가", null, "할인", null, null, null, "재고수량", null, "옵션", null, null, "상품이미지", null, null, null, null, null, null, null, null, null, null, null, null, null, "어린이제품", null, null, "전기용품", null, null, null, "생활용품", null, null, null, "생활화학/살생물제품", null, "원산지", null, null, null, "사은품/덤", null, null, null, null, null, null, null],
    [null, ...esmHeaders],
    [null, ...esmHeaders.map((_, index) => [0, 3, 9, 12, 13, 14, 19, 20, 24, 26, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 51, 52, 53, 54, 60, 61].includes(index) ? "필수" : "비필수")],
    [null, ...esmHeaders.map(() => "")],
    ["예시", "옥션/G마켓", "auctionid", "gmarketid", "상품명", "", "", "", "", 11111, 34470700, 100000077200002, "무제한", 100000, 100000, "", "", "", "", 99999, 99999, "미사용", "", "", "https://example.com/image.jpg", "", "<p>상세설명</p>", "", "일반택배", 175936, 663847, 440567, 399, 557, 10013, 2500, 7, 239479, "인증대상아님", "", "", "인증대상아님", "", "", "해당사항없음", "인증대상아님", "", "", "해당사항없음", "인증대상아님", "", "해당없음", "알수없음", "", "단일원산지", "", "", "", "", "", "구매가능", "과세상품", "가능"],
  ];
}

function makeEsmRow(product: Product, auctionId: string, gmarketId: string) {
  const row = Array(64).fill("");
  row[0] = "";
  row[1] = "옥션/G마켓";
  row[2] = auctionId;
  row[3] = gmarketId;
  row[4] = product.productName;
  row[10] = product.categoryCode;
  row[11] = product.auctionExposureCode;
  row[12] = product.gmarketExposureCode;
  row[13] = "무제한";
  row[14] = product.finalPrice;
  row[15] = product.finalPrice;
  row[20] = product.stock;
  row[21] = product.stock;
  row[22] = product.optionName ? "단독형" : "미사용";
  row[23] = product.optionName;
  row[24] = product.optionValue;
  row[25] = product.mainImage;
  row[26] = product.additionalImage;
  row[27] = product.detailHtml;
  row[29] = "일반택배";
  row[30] = product.departureCode;
  row[31] = product.shippingPolicyNumber;
  row[32] = product.returnAddressCode;
  row[33] = product.auctionShippingPolicy;
  row[34] = product.gmarketShippingPolicy;
  row[35] = product.courierCode;
  row[36] = product.returnShippingFee;
  row[37] = product.productGroupCode;
  row[38] = product.noticeTemplateCode;
  row[39] = "인증대상아님";
  row[42] = "인증대상아님";
  row[45] = "해당사항없음";
  row[46] = "인증대상아님";
  row[49] = "해당사항없음";
  row[50] = "인증대상아님";
  row[52] = product.originProductType;
  row[53] = product.originRegionType;
  row[54] = product.originRegionCode;
  row[55] = product.multipleOrigins;
  row[61] = "구매가능";
  row[62] = product.vatType.includes("면세") ? "면세상품" : "과세상품";
  row[63] = "가능";
  return row;
}

function makeEsmWorkbook(products: Product[], auctionId: string, gmarketId: string) {
  const sheet = XLSX.utils.aoa_to_sheet([...esmTopRows(), ...products.map(product => makeEsmRow(product, auctionId, gmarketId))]);
  sheet["!cols"] = Array(64).fill({ wch: 16 });
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet, "NEW 일반상품");
  return workbook;
}

function validateProducts(products: Product[], auctionId: string, gmarketId: string) {
  const issues: ValidationIssue[] = [];
  const requireField = (product: Product, value: unknown, field: string, step: StepId) => {
    if (normalize(value) === "" || value === 0) issues.push({ productId: product.id, productName: product.productName, field, message: `${field}이(가) 비어 있습니다.`, step });
  };
  products.forEach(product => {
    requireField(product, product.productName, "상품명", "category");
    requireField(product, product.categoryCode, "카테고리 코드", "category");
    if (auctionId) requireField(product, product.auctionExposureCode, "A 노출코드", "category");
    if (gmarketId) requireField(product, product.gmarketExposureCode, "G 노출코드", "category");
    requireField(product, product.productGroupCode, "상품군 코드", "category");
    requireField(product, product.noticeTemplateCode, "상품고시정보 템플릿코드", "notice");
    requireField(product, product.originProductType, "원산지 상품타입", "origin");
    requireField(product, product.originRegionType, "원산지 지역타입", "origin");
    requireField(product, product.departureCode, "출하지 코드", "shipping");
    requireField(product, product.shippingPolicyNumber, "배송정책번호", "shipping");
    requireField(product, product.returnAddressCode, "반품/교환 주소 코드", "shipping");
    if (auctionId) requireField(product, product.auctionShippingPolicy, "A 발송정책", "shipping");
    if (gmarketId) requireField(product, product.gmarketShippingPolicy, "G 발송정책", "shipping");
    requireField(product, product.courierCode, "택배사 코드", "shipping");
    requireField(product, product.returnShippingFee, "반품/교환 배송비", "price");
    requireField(product, product.finalPrice, "판매가", "price");
    requireField(product, product.stock, "재고", "price");
    requireField(product, product.mainImage, "기본이미지", "review");
    requireField(product, product.detailHtml, "상품상세설명", "review");
  });
  if (!auctionId && !gmarketId) issues.unshift({ productId: "ALL", productName: "전체", field: "판매자 ID", message: "옥션 또는 G마켓 판매자 ID를 입력해 주세요.", step: "review" });
  return issues;
}

const fieldStyle: React.CSSProperties = { display: "grid", gap: 6 };
const gridStyle: React.CSSProperties = { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(210px, 1fr))", gap: 12 };
const buttonStyle: React.CSSProperties = { padding: "10px 14px", borderRadius: 8, cursor: "pointer" };

export default function Home() {
  const [active, setActive] = useState<MarketTab>("smartstore");
  const [rows, setRows] = useState<Row[]>([]);
  const [products, setProducts] = useState<Product[]>([]);
  const [status, setStatus] = useState("상품 엑셀을 업로드해 주세요.");

  const [feeRate, setFeeRate] = useState(6);
  const [smartMargin, setSmartMargin] = useState(30);
  const [extraCost, setExtraCost] = useState(0);
  const [smartRound, setSmartRound] = useState(100);
  const [smartCategory, setSmartCategory] = useState("50001770");
  const [smartCourier, setSmartCourier] = useState("CJGLS");
  const [asPhone, setAsPhone] = useState("01027483227");

  const [step, setStep] = useState<StepId>("category");
  const [selectedGroup, setSelectedGroup] = useState("");
  const [auctionId, setAuctionId] = useState("");
  const [gmarketId, setGmarketId] = useState("");
  const [marginRate, setMarginRate] = useState(0);
  const [roundUnit, setRoundUnit] = useState(100);
  const [checking, setChecking] = useState(false);
  const [progress, setProgress] = useState(0);
  const [issues, setIssues] = useState<ValidationIssue[]>([]);
  const [checked, setChecked] = useState(false);

  const groups = useMemo(() => groupProducts(products, step), [products, step]);
  const currentGroup = groups.find(group => group.name === selectedGroup) ?? groups[0];

  async function readFile(file?: File) {
    if (!file) return;
    try {
      const workbook = XLSX.read(await file.arrayBuffer(), { type: "array" });
      const parsedRows = parseSheet(workbook.Sheets[workbook.SheetNames[0]]);
      const parsedProducts = makeProducts(parsedRows);
      setRows(parsedRows);
      setProducts(parsedProducts);
      setStep("category");
      setSelectedGroup("");
      setIssues([]);
      setChecked(false);
      setStatus(`${parsedProducts.length}개 상품을 읽었습니다. 카테고리 묶음부터 수정해 주세요.`);
    } catch {
      setStatus("파일을 읽지 못했습니다.");
    }
  }

  function patchGroup(groupName: string, patch: Partial<Product>) {
    const ids = new Set(groups.find(group => group.name === groupName)?.items.map(item => item.id) ?? []);
    setProducts(current => current.map(product => ids.has(product.id) ? { ...product, ...patch } : product));
    setChecked(false);
    setIssues([]);
    setStatus(`${groupName} 묶음 ${ids.size}개에 적용했습니다.`);
  }

  function moveNext() {
    const index = steps.findIndex(item => item.id === step);
    if (index < steps.length - 1) {
      setStep(steps[index + 1].id);
      setSelectedGroup("");
      setStatus(`${steps[index + 1].label} 단계로 이동했습니다.`);
    }
  }

  function downloadSmart() {
    const body = makeProducts(rows).map(product => [
      product.sellerCode,
      smartCategory,
      product.productName,
      "신상품",
      Math.ceil(((product.basePrice + extraCost) * (1 + (feeRate + smartMargin) / 100)) / smartRound) * smartRound,
      product.vatType || "과세상품",
      product.stock,
      product.mainImage,
      product.additionalImage,
      product.detailHtml,
      "03",
      "N",
      normalize(pick(product.raw, aliases.originDirect)),
      "Y",
      "택배, 소포, 등기",
      smartCourier,
      product.shippingFee > 0 ? "유료" : "무료",
      product.shippingFee,
      product.shippingFee || 3000,
      (product.shippingFee || 3000) * 2,
      "N",
      asPhone,
      "판매자에 문의하시거나, A/S연락처로 문의 주시기 바랍니다.",
      "Y",
      "N",
    ]);
    const sheet = XLSX.utils.aoa_to_sheet([naverHeaders, ...body]);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, sheet, "일괄등록");
    XLSX.writeFile(workbook, "postsheet02_스마트스토어.xlsx");
  }

  function applyPrices() {
    setProducts(current => current.map(product => ({ ...product, finalPrice: roundedPrice(product.basePrice, marginRate, roundUnit) })));
    setChecked(false);
    setIssues([]);
    setStatus(`전체 ${products.length}개 상품에 마진율 ${marginRate}%와 올림 단위 ${roundUnit}원을 적용했습니다.`);
  }

  async function runReview() {
    setChecking(true);
    setChecked(false);
    setProgress(5);
    setStatus("ESM 양식과 필수항목을 검사하고 있습니다.");
    await new Promise(resolve => setTimeout(resolve, 150));
    setProgress(35);
    await new Promise(resolve => setTimeout(resolve, 150));
    const found = validateProducts(products, auctionId.trim(), gmarketId.trim());
    setProgress(75);
    await new Promise(resolve => setTimeout(resolve, 150));
    setIssues(found);
    setProgress(100);
    setChecked(true);
    setChecking(false);
    setStatus(found.length === 0 ? `최종검사 완료: ${products.length}개 상품이 ESM 필수형식을 통과했습니다.` : `최종검사 실패: ${found.length}건을 수정해야 합니다.`);
  }

  async function downloadReviewed() {
    if (!checked || issues.length > 0) return;
    const byCategory = new Map<string, Product[]>();
    products.forEach(product => {
      const key = product.categoryCode || "미분류";
      byCategory.set(key, [...(byCategory.get(key) ?? []), product]);
    });
    const zip = new JSZip();
    for (const [category, items] of byCategory.entries()) {
      for (let index = 0; index < items.length; index += 500) {
        const chunk = items.slice(index, index + 500);
        const suffix = items.length > 500 ? `_${String(index / 500 + 1).padStart(2, "0")}` : "";
        const name = `ESM_카테고리_${category}${suffix}_${chunk.length}개.xlsx`;
        zip.file(name, XLSX.write(makeEsmWorkbook(chunk, auctionId.trim(), gmarketId.trim()), { bookType: "xlsx", type: "array" }));
      }
    }
    const blob = await zip.generateAsync({ type: "blob" });
    const anchor = document.createElement("a");
    anchor.href = URL.createObjectURL(blob);
    anchor.download = `ESM_검수완료_${products.length}개_분할파일.zip`;
    anchor.click();
    URL.revokeObjectURL(anchor.href);
  }

  function renderGroupEditor() {
    if (!currentGroup) return <p>수정할 상품이 없습니다.</p>;
    const groupName = currentGroup.name;
    const sample = currentGroup.items.slice(0, 8);

    if (step === "category") {
      const first = currentGroup.items[0];
      return <GroupCard title={groupName} count={currentGroup.items.length} sample={sample}>
        <div style={gridStyle}>
          <Input label="카테고리 코드" defaultValue={first.categoryCode} onApply={value => patchGroup(groupName, { categoryCode: digits(value), categoryGroup: `카테고리 ${digits(value) || "미입력"}` })} />
          <Input label="A 노출코드" defaultValue={first.auctionExposureCode} onApply={value => patchGroup(groupName, { auctionExposureCode: digits(value) })} />
          <Input label="G 노출코드" defaultValue={first.gmarketExposureCode} onApply={value => patchGroup(groupName, { gmarketExposureCode: digits(value) })} />
          <Input label="상품군 코드" defaultValue={first.productGroupCode} onApply={value => patchGroup(groupName, { productGroupCode: digits(value) })} />
        </div>
      </GroupCard>;
    }
    if (step === "notice") {
      const first = currentGroup.items[0];
      return <GroupCard title={groupName} count={currentGroup.items.length} sample={sample}>
        <Input label="상품고시정보 템플릿코드" defaultValue={first.noticeTemplateCode || "239479"} onApply={value => patchGroup(groupName, { noticeTemplateCode: digits(value) || "239479" })} />
      </GroupCard>;
    }
    if (step === "origin") {
      const first = currentGroup.items[0];
      return <GroupCard title={groupName} count={currentGroup.items.length} sample={sample}>
        <div style={gridStyle}>
          <Input label="원산지 상품타입" defaultValue={first.originProductType} onApply={value => patchGroup(groupName, { originProductType: value })} />
          <Input label="원산지 지역타입" defaultValue={first.originRegionType} onApply={value => patchGroup(groupName, { originRegionType: value })} />
          <Input label="원산지 지역코드" defaultValue={first.originRegionCode} onApply={value => patchGroup(groupName, { originRegionCode: digits(value), originGroup: value || "원산지 미확인" })} />
          <Input label="복수 원산지여부" defaultValue={first.multipleOrigins} onApply={value => patchGroup(groupName, { multipleOrigins: value })} />
        </div>
      </GroupCard>;
    }
    if (step === "shipping") {
      const first = currentGroup.items[0];
      return <GroupCard title={groupName} count={currentGroup.items.length} sample={sample}>
        <div style={gridStyle}>
          <Input label="출하지 코드" defaultValue={first.departureCode} onApply={value => patchGroup(groupName, { departureCode: digits(value) })} />
          <Input label="배송정책번호" defaultValue={first.shippingPolicyNumber} onApply={value => patchGroup(groupName, { shippingPolicyNumber: digits(value) })} />
          <Input label="반품/교환 주소 코드" defaultValue={first.returnAddressCode} onApply={value => patchGroup(groupName, { returnAddressCode: digits(value) })} />
          <Input label="A 발송정책" defaultValue={first.auctionShippingPolicy} onApply={value => patchGroup(groupName, { auctionShippingPolicy: digits(value) })} />
          <Input label="G 발송정책" defaultValue={first.gmarketShippingPolicy} onApply={value => patchGroup(groupName, { gmarketShippingPolicy: digits(value) })} />
          <Input label="택배사 코드" defaultValue={first.courierCode} onApply={value => patchGroup(groupName, { courierCode: digits(value) })} />
        </div>
      </GroupCard>;
    }
    return null;
  }

  return <main className="container">
    <section className="hero">
      <span className="badge">postsheet02 · 상품 대량등록 변환</span>
      <h1>상품 일괄목록을<br />마켓 등록 파일로 변환</h1>
    </section>

    <section className="panel">
      <div className="field full">
        <label>1. 상품 일괄목록 엑셀 업로드</label>
        <input type="file" accept=".xlsx,.xls,.csv" onChange={event => readFile(event.target.files?.[0])} />
      </div>
    </section>

    <section className="panel">
      <div className="full" style={{ display: "flex", gap: 8, marginBottom: 16, flexWrap: "wrap" }}>
        <button style={buttonStyle} onClick={() => setActive("smartstore")}>스마트스토어</button>
        <button style={buttonStyle} onClick={() => setActive("esm")}>ESM · 옥션/G마켓</button>
      </div>

      {active === "smartstore" ? <>
        <div style={gridStyle} className="full">
          <label style={fieldStyle}>네이버 수수료율 (%)<input type="number" value={feeRate} onChange={event => setFeeRate(Number(event.target.value))} /></label>
          <label style={fieldStyle}>추가 마진율 (%)<input type="number" value={smartMargin} onChange={event => setSmartMargin(Number(event.target.value))} /></label>
          <label style={fieldStyle}>상품당 추가비용<input type="number" value={extraCost} onChange={event => setExtraCost(Number(event.target.value))} /></label>
          <label style={fieldStyle}>판매가 올림 단위<input type="number" value={smartRound} onChange={event => setSmartRound(Number(event.target.value))} /></label>
          <label style={fieldStyle}>카테고리코드<input value={smartCategory} onChange={event => setSmartCategory(event.target.value)} /></label>
          <label style={fieldStyle}>택배사코드<input value={smartCourier} onChange={event => setSmartCourier(event.target.value)} /></label>
          <label style={fieldStyle}>A/S 전화번호<input value={asPhone} onChange={event => setAsPhone(event.target.value)} /></label>
        </div>
        <div className="actions full"><button style={buttonStyle} disabled={rows.length === 0} onClick={downloadSmart}>스마트스토어 다운로드</button></div>
      </> : <>
        <div className="full" style={{ display: "flex", gap: 8, overflowX: "auto", marginBottom: 18 }}>
          {steps.map(item => <button key={item.id} style={{ ...buttonStyle, whiteSpace: "nowrap", fontWeight: step === item.id ? 700 : 400 }} onClick={() => { setStep(item.id); setSelectedGroup(""); }}>{item.label}</button>)}
        </div>

        <div className="full" style={gridStyle}>
          <label style={fieldStyle}>옥션 판매자 ID<input value={auctionId} onChange={event => { setAuctionId(event.target.value); setChecked(false); }} /></label>
          <label style={fieldStyle}>G마켓 판매자 ID<input value={gmarketId} onChange={event => { setGmarketId(event.target.value); setChecked(false); }} /></label>
        </div>

        {products.length > 0 && ["category", "notice", "origin", "shipping"].includes(step) && <div className="full" style={{ display: "grid", gridTemplateColumns: "minmax(220px, 300px) 1fr", gap: 16, marginTop: 18 }}>
          <aside style={{ display: "grid", gap: 8, alignContent: "start" }}>
            <strong>{groups.length}개 묶음 · {products.length}개 상품</strong>
            {groups.map(group => <button key={group.name} style={{ ...buttonStyle, textAlign: "left" }} onClick={() => setSelectedGroup(group.name)}>{group.name}<br /><small>{group.items.length}개</small></button>)}
          </aside>
          <div>{renderGroupEditor()}</div>
        </div>}

        {step === "price" && <div className="full" style={{ marginTop: 18 }}>
          <h2>가격과 배송비는 마지막에 적용합니다</h2>
          <div style={gridStyle}>
            <label style={fieldStyle}>마진율 (%)<input type="number" value={marginRate} onChange={event => setMarginRate(Number(event.target.value))} /></label>
            <label style={fieldStyle}>판매가 올림 단위<input type="number" value={roundUnit} onChange={event => setRoundUnit(Number(event.target.value))} /></label>
          </div>
          <button style={{ ...buttonStyle, marginTop: 12 }} onClick={applyPrices}>전체 상품 가격 적용</button>
          <p>마진율 0%이면 원래 가격을 그대로 사용합니다.</p>
        </div>}

        {step === "review" && <div className="full" style={{ marginTop: 18 }}>
          <h2>ESM 최종검사</h2>
          <p>ESM 열 구조, 필수항목 누락, 마켓별 A/G 필수값을 검사합니다.</p>
          <button style={buttonStyle} disabled={checking || products.length === 0} onClick={runReview}>{checking ? "검사 중…" : "최종검사 시작"}</button>
          {checking && <div style={{ marginTop: 14 }}><progress value={progress} max={100} style={{ width: "100%" }} /><p>{progress}% · ESM 등록파일을 검사하고 있습니다.</p></div>}
          {checked && issues.length === 0 && <div style={{ marginTop: 14 }}><h3>검사 완료</h3><p>필수항목 누락이 없습니다. 카테고리별, 최대 500개 단위로 분리해 ZIP으로 다운로드합니다.</p><button style={buttonStyle} onClick={downloadReviewed}>검수 완료 파일 다운로드</button></div>}
          {checked && issues.length > 0 && <div style={{ marginTop: 14 }}><h3>수정 필요: {issues.length}건</h3>{Object.entries(issues.reduce<Record<string, number>>((acc, issue) => ({ ...acc, [issue.field]: (acc[issue.field] ?? 0) + 1 }), {})).map(([field, count]) => <button key={field} style={{ ...buttonStyle, margin: 4 }} onClick={() => { const target = issues.find(issue => issue.field === field); if (target) setStep(target.step); }}>{field} {count}건</button>)}<div style={{ maxHeight: 280, overflow: "auto", marginTop: 10 }}>{issues.slice(0, 100).map((issue, index) => <p key={`${issue.productId}-${issue.field}-${index}`}>{issue.productName} · {issue.message}</p>)}</div></div>}
        </div>}

        {step !== "review" && products.length > 0 && <div className="actions full" style={{ marginTop: 18 }}><button style={buttonStyle} onClick={moveNext}>저장하고 다음 단계</button></div>}
      </>}

      <div className="status full">{status}</div>
    </section>
  </main>;
}

function Input({ label, defaultValue, onApply }: { label: string; defaultValue: string | number; onApply: (value: string) => void }) {
  const [value, setValue] = useState(String(defaultValue ?? ""));
  return <label style={fieldStyle}>{label}<input value={value} onChange={event => setValue(event.target.value)} /><button type="button" style={buttonStyle} onClick={() => onApply(value)}>이 묶음에 적용</button></label>;
}

function GroupCard({ title, count, sample, children }: { title: string; count: number; sample: Product[]; children: React.ReactNode }) {
  return <div style={{ border: "1px solid #ddd", borderRadius: 12, padding: 16 }}>
    <h2>{title}</h2>
    <p>{count}개 상품을 한 번에 수정합니다.</p>
    <details style={{ marginBottom: 16 }}><summary>상품명 확인</summary>{sample.map(product => <p key={product.id}>{product.productName}</p>)}</details>
    {children}
  </div>;
}

"use client";

import { useMemo, useState } from "react";
import * as XLSX from "xlsx";
import JSZip from "jszip";

type Row = Record<string, unknown>;
type MarketTab = "smartstore" | "esm";

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
  courierName: "CJ대한통운" | "한진택배";
};

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
  stock: ["재고수량", "재고", "수량", "판매가능수량", "A 재고", "G 재고"],
  optionName: ["옵션명", "옵션항목", "옵션"],
  optionValue: ["옵션값", "옵션내용", "선택옵션", "옵션 입력값"],
  mainImage: ["대표이미지", "대표이미지URL", "이미지1", "메인이미지", "이미지URL", "기본이미지"],
  detailHtml: ["상세설명", "상품설명", "상세HTML", "상품상세", "상세페이지", "상세이미지", "상세이미미", "상품상세설명"],
  shippingFee: ["기본배송비", "배송비", "반품/교환 배송비"],
  brand: ["브랜드"], maker: ["제조사", "제조자"],
  vatType: ["부가세", "과세구분", "부가세유형", "부가세여부"],
  originCode: ["원산지코드", "원산지 지역코드"],
  originDirect: ["원산지 직접입력", "원산지직접입력", "원산지", "원산지명"],
};

const naverHeaders = [
  "판매자 상품코드","카테고리코드","상품명","상품상태","판매가","단위가격 사용여부","표시용량","표시단위","총용량","부가세","관부가세","재고수량","옵션형태","옵션명","옵션값","옵션가","옵션 재고수량","직접입력 옵션","추가상품명","추가상품값","추가상품가","추가상품 재고수량","대표이미지","추가이미지","상세설명","브랜드","제조사","제조일자","유효일자","원산지코드","수입사","복수원산지여부","원산지 직접입력","미성년자 구매","배송비 템플릿코드","배송방법","택배사코드","배송비유형","기본배송비","배송비 결제방식","조건부무료-\n상품판매가 합계","수량별부과-수량","구간별-\n2구간수량","구간별-\n3구간수량","구간별-\n3구간배송비","구간별-\n추가배송비","반품배송비","교환배송비","지역별 차등 배송비","별도설치비","상품정보제공고시 템플릿코드","상품정보제공고시\n품명","상품정보제공고시\n모델명","상품정보제공고시\n인증허가사항","상품정보제공고시\n제조자","A/S 템플릿코드","A/S 전화번호","A/S 안내","판매자특이사항","즉시할인 값\n(기본할인)","즉시할인 단위\n(기본할인)","모바일\n즉시할인 값","모바일\n즉시할인 단위","복수구매할인\n조건 값","복수구매할인\n조건 단위","복수구매할인\n값","복수구매할인\n단위","상품구매시 포인트\n지급 값","상품구매시 포인트\n지급 단위","텍스트리뷰 작성시\n지급 포인트","포토/동영상 리뷰 작성시\n지급 포인트","한달사용 텍스트리뷰\n작성시 지급 포인트","한달사용\n포토/동영상리뷰 작성시 지급 포인트","알림받기동의 고객 리뷰 작성 시 지급 포인트","무이자\n할부 개월","사은품","판매자바코드","구매평 노출여부","구매평\n비노출사유","알림받기 동의 고객 전용 여부","ISBN","ISSN","독립출판","출간일","출판사","글작가","그림작가","번역자명","문화비 소득공제","사이즈\n상품군","사이즈\n사이즈명","사이즈\n상세 사이즈","사이즈\n모델명","판매준수가"
];
const naverGroupRow = naverHeaders.map((_, i) => i===0?"상품 기본정보":i===25?"상품 주요정보":i===34?"배송정보":i===50?"상품정보제공고시":i===55?"A/S, 특이사항":i===59?"할인/혜택정보":i===77?"기타 정보":"");

const esmHeaders = ["노출\n사이트","A ID","G ID","상품명","A프로모션 문구","G프로모션 문구","G 영문","G 중문","카테고리 템플릿 코드","카테고리 코드","A 노출코드","G 노출코드","판매기간","A 판매가","G 판매가","A 할인유형","A 할인가","G 할인유형","G 할인가","A 재고","G 재고","옵션\n타입","옵션명","옵션\n입력값","기본이미지","추가이미지","상품상세설명","배송정보 \n템플릿 코드","배송방법","출하지 코드","배송정책번호","반품/교환\n주소 코드","A 발송정책","G 발송정책","택배사\n코드","반품/교환\n배송비","상품군\n코드","상품고시정보\n템플릿코드","인증타입","인증품목선택","인증코드","인증타입","인증품목선택","인증코드","병행수입여부","인증타입","인증품목선택","인증코드","병행수입여부","인증타입","승인/신고번호","원산지\n상품타입","원산지\n지역타입","원산지\n지역코드","복수\n원산지여부","사은품/덤 \n템플릿 코드","사은품","덤","소비기한","제조일자","청소년구매\n불가여부","부가세여부","선물하기상품"];
const esmTopRows: unknown[][] = [
  ["▶가이드 바로가기",null," ※ 문서버전 : NEW 2.0",null,"필독▶ 상품정보는 8행부터 입력합니다. 1~7행은 삭제하지 마세요."],
  [null,"상품기본정보",...Array(26).fill(null),"배송정보",...Array(8).fill(null),"상품고시정보",...Array(25).fill(null),"추가정보"],
  [null,"계정선택",null,null,"상품명",null,null,null,null,"카테고리",null,null,null,"판매가",null,"할인",null,null,null,"재고수량",null," 옵션",null,null,"상품이미지",null,null,null,null,"배송정보"],
  [null,...esmHeaders],
  [null,...esmHeaders.map((h,i)=>[0,1,2,3,9,10,11,12,13,14,19,20,24,26,28,29,30,31,32,33,34,35,36,37,51,52,53,54,60,61].includes(i)?"필수":"비필수")],
  [null,...esmHeaders.map(()=>"")],
  ["예시","옥션/G마켓","auctionid","gmarketid","상품명","","","","","","","","","무제한",100000,100000,"","","","",99999,99999,"미사용","","","https://example.com/1.jpg","https://example.com/2.jpg","<p>상품설명</p>","","일반택배","","","","","",10013,2500,"","","인증대상아님","","","인증대상아님","","","해당사항없음","인증대상아님","","","해당사항없음","인증대상아님","","해당없음","알수없음","","단일원산지","","","","","","구매가능","과세상품","가능"]
];

function cleanKey(v: unknown){return String(v??"").replace(/[\s\n\r_\-()\[\]\/]/g,"").toLowerCase();}
function normalize(v: unknown){return String(v??"").trim();}
function numberValue(v: unknown){const n=Number(String(v??"").replace(/,/g,"").replace(/[^0-9.-]/g,""));return Number.isFinite(n)?n:0;}
function pickExact(row:Row,name:string){const t=cleanKey(name);const f=Object.entries(row).find(([k])=>cleanKey(k)===t);return f?f[1]:"";}
function pick(row:Row,names:string[]){for(const n of names){const v=pickExact(row,n);if(normalize(v)!=="")return v;}const e=Object.entries(row);for(const n of names){const t=cleanKey(n);const f=e.find(([k])=>cleanKey(k).includes(t)||t.includes(cleanKey(k)));if(f)return f[1];}return "";}
function getAdditionalImage(row:Row){return normalize(pickExact(row,"이미지2"))||normalize(pickExact(row,"추가이미지"));}
function parseSheet(sheet:XLSX.WorkSheet){const matrix=XLSX.utils.sheet_to_json<unknown[]>(sheet,{header:1,defval:"",raw:false});const candidates=[...aliases.productName,...aliases.basePrice,...aliases.sellerCode,...aliases.stock].map(cleanKey);let best=0,score=-1;matrix.slice(0,30).forEach((r,i)=>{const s=r.map(cleanKey).filter(Boolean).reduce((a,k)=>a+(candidates.some(c=>k===c||k.includes(c)||c.includes(k))?1:0),0);if(s>score){score=s;best=i;}});const raw=(matrix[best]??[]).map((v,i)=>normalize(v)||`열${i+1}`);const seen=new Map<string,number>();const headers=raw.map(h=>{const c=(seen.get(h)||0)+1;seen.set(h,c);return c===1?h:`${h}_${c}`;});return matrix.slice(best+1).map(v=>Object.fromEntries(headers.map((h,i)=>[h,v[i]??""]))).filter(r=>Object.values(r).some(v=>normalize(v)!==""));}
function toProducts(rows:Row[]):Product[]{return rows.map(raw=>({raw,productName:normalize(pick(raw,aliases.productName)),sellerCode:normalize(pick(raw,aliases.sellerCode)),basePrice:numberValue(pick(raw,aliases.basePrice)),stock:numberValue(pick(raw,aliases.stock))||99999,optionName:normalize(pick(raw,aliases.optionName)),optionValue:normalize(pick(raw,aliases.optionValue)),mainImage:normalize(pick(raw,aliases.mainImage)),additionalImage:getAdditionalImage(raw),detailHtml:normalize(pick(raw,aliases.detailHtml)),shippingFee:numberValue(pick(raw,aliases.shippingFee)),brand:normalize(pick(raw,aliases.brand)),maker:normalize(pick(raw,aliases.maker)),vatType:normalize(pick(raw,aliases.vatType)),originCode:normalize(pick(raw,aliases.originCode)),originDirect:normalize(pick(raw,aliases.originDirect))})).filter(p=>p.productName||p.sellerCode);}
function roundedPrice(base:number,margin:number,unit:number,extra=0){if(base<=0)return 0;return Math.ceil(((base+extra)*(1+margin/100))/unit)*unit;}
function downloadBlob(blob:Blob,name:string){const url=URL.createObjectURL(blob);const a=document.createElement("a");a.href=url;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);}
function copyCode(raw:Row,names:string[]){return normalize(pick(raw,names));}
function courierCode(name:EsmSettings["courierName"]){return name==="CJ대한통운"?10013:10007;}
function esmOptionType(p:Product){if(!p.optionName||!p.optionValue)return "미사용";const count=p.optionName.split(/[,\n]/).filter(Boolean).length;return count>=3?"3개조합형":count===2?"2개조합형":"단독형";}
function esmRow(p:Product,s:EsmSettings){const raw=p.raw;const price=roundedPrice(p.basePrice,s.marginRate,s.roundUnit);const row=Array(esmHeaders.length).fill("");const set=(h:string,v:unknown)=>{const i=esmHeaders.indexOf(h);if(i>=0)row[i]=v??"";};set("노출\n사이트","옥션/G마켓");set("A ID",s.auctionId);set("G ID",s.gmarketId);set("상품명",p.productName);set("카테고리 템플릿 코드",copyCode(raw,["카테고리 템플릿 코드","카테고리템플릿코드"]));set("카테고리 코드",copyCode(raw,["ESM 카테고리코드","카테고리 코드","카테고리코드"]));set("A 노출코드",copyCode(raw,["A 노출코드","옥션 노출코드"]));set("G 노출코드",copyCode(raw,["G 노출코드","G마켓 노출코드"]));set("판매기간","무제한");set("A 판매가",price);set("G 판매가",price);set("A 재고",p.stock);set("G 재고",p.stock);set("옵션\n타입",esmOptionType(p));set("옵션명",p.optionName);set("옵션\n입력값",p.optionValue);set("기본이미지",p.mainImage);set("추가이미지",p.additionalImage);set("상품상세설명",p.detailHtml);set("배송정보 \n템플릿 코드",copyCode(raw,["배송정보 템플릿 코드","배송정보템플릿코드"]));set("배송방법","일반택배");set("출하지 코드",copyCode(raw,["출하지 코드","출하지코드"]));set("배송정책번호",copyCode(raw,["배송정책번호"]));set("반품/교환\n주소 코드",copyCode(raw,["반품/교환 주소 코드","반품교환주소코드"]));set("A 발송정책",copyCode(raw,["A 발송정책","옥션 발송정책"]));set("G 발송정책",copyCode(raw,["G 발송정책","G마켓 발송정책"]));set("택배사\n코드",courierCode(s.courierName));set("반품/교환\n배송비",numberValue(pick(raw,["반품/교환 배송비","반품배송비","교환배송비"]))||p.shippingFee||2500);set("상품군\n코드",copyCode(raw,["상품군 코드","상품군코드"]));set("상품고시정보\n템플릿코드",copyCode(raw,["상품고시정보 템플릿코드","상품고시 템플릿코드"]));set("인증타입","인증대상아님");set("병행수입여부","해당사항없음");set("원산지\n상품타입",copyCode(raw,["원산지 상품타입"])||"해당없음");set("원산지\n지역타입",copyCode(raw,["원산지 지역타입"])||"알수없음");set("원산지\n지역코드",p.originCode);set("복수\n원산지여부",copyCode(raw,["복수원산지여부"])||"단일원산지");set("청소년구매\n불가여부","구매가능");set("부가세여부",p.vatType.includes("면세")?"면세상품":"과세상품");set("선물하기상품","가능");return ["",...row];}
function makeEsmWorkbook(rows:Product[],s:EsmSettings){const data=[...esmTopRows,...rows.map(p=>esmRow(p,s))];const sheet=XLSX.utils.aoa_to_sheet(data);sheet["!cols"]=Array(64).fill({wch:16});sheet["!rows"]=[{hpt:24},{hpt:24},{hpt:24},{hpt:50},{hpt:24},{hpt:120},{hpt:32}];const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,sheet,"NEW 일반상품");return wb;}

export default function Home(){
  const [active,setActive]=useState<MarketTab>("smartstore");
  const [sourceRows,setSourceRows]=useState<Row[]>([]);const [fileName,setFileName]=useState("");const [status,setStatus]=useState("상품 엑셀을 업로드해 주세요.");
  const products=useMemo(()=>toProducts(sourceRows),[sourceRows]);
  const [feeRate,setFeeRate]=useState(6),[marginRate,setMarginRate]=useState(30),[extraCost,setExtraCost]=useState(0),[roundUnit,setRoundUnit]=useState(100),[categoryCode,setCategoryCode]=useState("50001770"),[naverCourier,setNaverCourier]=useState("CJGLS"),[asPhone,setAsPhone]=useState("01027483227"),[multipleOrigins,setMultipleOrigins]=useState<"N"|"Y">("N");
  const [smartApplied,setSmartApplied]=useState<SmartSettings|null>(null);
  const [esmMargin,setEsmMargin]=useState(30),[esmRound,setEsmRound]=useState(100),[auctionId,setAuctionId]=useState(""),[gmarketId,setGmarketId]=useState(""),[esmCourier,setEsmCourier]=useState<EsmSettings["courierName"]>("CJ대한통운");
  const [esmApplied,setEsmApplied]=useState<EsmSettings|null>(null);

  async function readFile(file?:File){if(!file)return;setFileName(file.name);setSmartApplied(null);setEsmApplied(null);try{const wb=XLSX.read(await file.arrayBuffer(),{type:"array"});const rows=parseSheet(wb.Sheets[wb.SheetNames[0]]);setSourceRows(rows);setStatus(`${rows.length}행을 읽었습니다. 마켓 탭에서 설정을 적용해 주세요.`);}catch{setSourceRows([]);setStatus("파일을 읽지 못했습니다.");}}
  function applySmart(){if(!products.length)return;setSmartApplied({feeRate,marginRate,extraCost,roundUnit,categoryCode:categoryCode.trim(),courierCode:naverCourier.trim(),asPhone:asPhone.trim(),multipleOrigins});setStatus(`스마트스토어 ${products.length}개 적용 완료`);}
  function smartRow(p:Product,s:SmartSettings){const row=naverHeaders.map(h=>pickExact(p.raw,h));const set=(h:string,v:unknown)=>{const i=naverHeaders.indexOf(h);if(i>=0)row[i]=v??"";};set("판매자 상품코드",p.sellerCode);set("카테고리코드",s.categoryCode);set("상품명",p.productName);set("상품상태","신상품");set("판매가",roundedPrice(p.basePrice,s.marginRate+s.feeRate,s.roundUnit,s.extraCost));set("단위가격 사용여부","N");set("부가세",p.vatType||"과세상품");set("재고수량",p.stock===99999?1:p.stock);set("대표이미지",p.mainImage);set("추가이미지",p.additionalImage);set("상세설명",p.detailHtml);set("원산지코드","03");set("복수원산지여부",s.multipleOrigins);set("원산지 직접입력",p.originDirect);set("미성년자 구매","Y");set("배송방법","택배, 소포, 등기");set("택배사코드",s.courierCode);set("배송비유형",p.shippingFee>0?"유료":"무료");set("기본배송비",p.shippingFee);set("반품배송비",p.shippingFee||3000);set("교환배송비",(p.shippingFee||3000)*2);set("별도설치비","N");set("A/S 전화번호",s.asPhone);set("A/S 안내","판매자에 문의하시거나, A/S연락처로 문의 주시기 바랍니다.");set("구매평 노출여부","Y");set("알림받기 동의 고객 전용 여부","N");return row;}
  function downloadSmart(){if(!smartApplied)return;const sheet=XLSX.utils.aoa_to_sheet([naverGroupRow,naverHeaders,...products.map(p=>smartRow(p,smartApplied))]);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,sheet,"일괄등록");XLSX.writeFile(wb,`postsheet02_스마트스토어_${new Date().toISOString().slice(0,10)}.xlsx`);}
  function applyEsm(){if(!products.length)return;if(!auctionId.trim()&&!gmarketId.trim()){setStatus("옥션 또는 G마켓 판매자 ID를 하나 이상 입력해 주세요.");return;}const s={marginRate:esmMargin,roundUnit:esmRound,auctionId:auctionId.trim(),gmarketId:gmarketId.trim(),courierName:esmCourier};setEsmApplied(s);setStatus(`ESM ${products.length}개 적용 완료 · ${Math.ceil(products.length/500)}개 파일로 자동 분할됩니다.`);}
  async function downloadEsm(){if(!esmApplied)return;const chunks:Product[][]=[];for(let i=0;i<products.length;i+=500)chunks.push(products.slice(i,i+500));const date=new Date().toISOString().slice(0,10);if(chunks.length===1){XLSX.writeFile(makeEsmWorkbook(chunks[0],esmApplied),`ESM_${chunks[0].length}개_${date}.xlsx`);return;}const zip=new JSZip();chunks.forEach((c,i)=>{const out=XLSX.write(makeEsmWorkbook(c,esmApplied),{bookType:"xlsx",type:"array"});zip.file(`ESM_${String(i+1).padStart(2,"0")}_${c.length}개.xlsx`,out);});downloadBlob(await zip.generateAsync({type:"blob"}),`ESM_${products.length}개_분할파일_${date}.zip`);}

  return <main className="container"><section className="hero"><span className="badge">postsheet02 · 상품 대량등록 변환</span><h1>상품 일괄목록을<br/>마켓 등록 파일로 변환</h1><p>업로드 파일을 자동 해석해 스마트스토어와 ESM 양식으로 변환합니다.</p><div className="privacy">원본·결과 파일 서버 저장 없음 · 브라우저 안에서만 처리</div></section>
  <section className="panel"><div className="field full"><label>1. 상품 일괄목록 엑셀 업로드</label><input type="file" accept=".xlsx,.xls,.csv" onChange={e=>readFile(e.target.files?.[0])}/><small>{fileName||"선택된 파일 없음"}</small></div></section>
  <section className="panel"><div className="full" style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:16}}><button onClick={()=>setActive("smartstore")} style={{opacity:active==="smartstore"?1:.6}}>스마트스토어</button><button onClick={()=>setActive("esm")} style={{opacity:active==="esm"?1:.6}}>ESM · 옥션/G마켓</button></div>
  {active==="smartstore"&&<><div className="field"><label>네이버 수수료율 (%)</label><input type="number" value={feeRate} onChange={e=>setFeeRate(+e.target.value)}/></div><div className="field"><label>추가 마진율 (%)</label><input type="number" value={marginRate} onChange={e=>setMarginRate(+e.target.value)}/></div><div className="field"><label>상품당 추가비용</label><input type="number" value={extraCost} onChange={e=>setExtraCost(+e.target.value)}/></div><div className="field"><label>판매가 올림 단위</label><select value={roundUnit} onChange={e=>setRoundUnit(+e.target.value)}><option value={1}>1원</option><option value={10}>10원</option><option value={100}>100원</option><option value={500}>500원</option><option value={1000}>1,000원</option></select></div><div className="field"><label>카테고리코드</label><input value={categoryCode} onChange={e=>setCategoryCode(e.target.value)}/></div><div className="field"><label>택배사코드</label><input value={naverCourier} onChange={e=>setNaverCourier(e.target.value)}/></div><div className="field"><label>A/S 전화번호</label><input value={asPhone} onChange={e=>setAsPhone(e.target.value)}/></div><div className="field"><label>복수원산지</label><select value={multipleOrigins} onChange={e=>setMultipleOrigins(e.target.value as "N"|"Y")}><option value="N">단일</option><option value="Y">복수</option></select></div><div className="actions full"><button onClick={downloadSmart} disabled={!smartApplied}>스마트스토어 다운로드</button><button onClick={applySmart} disabled={!products.length}>위 내용 적용하기</button></div></>}
  {active==="esm"&&<><div className="field full"><label>ESM 자동 변환</label><small>상품명·가격·재고·이미지·상세설명·옵션·부가세·원산지와 원본에 있는 ESM 코드를 자동 연결합니다.</small></div><div className="field"><label>마진율 (%)</label><input type="number" min={0} value={esmMargin} onChange={e=>setEsmMargin(+e.target.value)}/></div><div className="field"><label>판매가 올림 단위</label><select value={esmRound} onChange={e=>setEsmRound(+e.target.value)}><option value={1}>1원</option><option value={10}>10원</option><option value={100}>100원</option><option value={500}>500원</option><option value={1000}>1,000원</option></select></div><div className="field"><label>옥션 판매자 ID</label><input value={auctionId} onChange={e=>setAuctionId(e.target.value)}/></div><div className="field"><label>G마켓 판매자 ID</label><input value={gmarketId} onChange={e=>setGmarketId(e.target.value)}/></div><div className="field"><label>발송 택배사</label><select value={esmCourier} onChange={e=>setEsmCourier(e.target.value as EsmSettings["courierName"])}><option value="CJ대한통운">CJ대한통운 · 10013</option><option value="한진택배">한진택배 · 10007</option></select></div><div className="field full"><small>{products.length>500?`${products.length}개 상품을 500개 단위로 ${Math.ceil(products.length/500)}개 엑셀에 나누어 ZIP으로 다운로드합니다.`:"500개 이하이므로 엑셀 한 개로 다운로드합니다."}</small></div><div className="actions full"><button onClick={downloadEsm} disabled={!esmApplied}>ESM 파일 다운로드</button><button onClick={applyEsm} disabled={!products.length}>위 내용 적용하기</button></div></>}
  <div className="status full">{status}</div></section>
  <section className="preview"><div className="previewHead"><h2>자동 해석 미리보기</h2><span>{products.length}개 상품</span></div><div className="tableWrap"><table><thead><tr><th>상품명</th><th>원래 판매가</th><th>재고</th><th>이미지1</th><th>이미지2</th><th>상품설명</th></tr></thead><tbody>{products.slice(0,30).map((p,i)=><tr key={`${p.sellerCode}-${i}`}><td>{p.productName}</td><td>{p.basePrice.toLocaleString()}</td><td>{p.stock}</td><td>{p.mainImage?"있음":"없음"}</td><td>{p.additionalImage?"있음":"없음"}</td><td>{p.detailHtml?"있음":"없음"}</td></tr>)}{!products.length&&<tr><td colSpan={6} className="empty">파일을 업로드해 주세요.</td></tr>}</tbody></table></div></section></main>;
}

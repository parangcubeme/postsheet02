"use client";

import { useMemo, useState } from "react";
import * as XLSX from "xlsx";
import JSZip from "jszip";

type Row = Record<string, unknown>;
type Courier = "CJ대한통운" | "한진택배";
type Product = { raw: Row; name: string; code: string; price: number; stock: number; image1: string; image2: string; detail: string; optionName: string; optionValue: string; shipping: number; origin: string; vat: string };

const aliases = {
  name:["상품명","제품명","품명"], code:["판매자 상품코드","판매자상품코드","상품코드","관리코드"],
  price:["판매가","판매가격","상품가격","공급가","공급가격","단가","원가","매입가","가격"], stock:["재고수량","재고","수량"],
  image1:["이미지1","대표이미지","대표이미지URL","기본이미지"], image2:["이미지2","추가이미지"], detail:["상품설명","상세설명","상품상세설명"],
  optionName:["옵션명","옵션"], optionValue:["옵션값","옵션내용"], shipping:["기본배송비","배송비"], origin:["원산지코드","원산지 지역코드"], vat:["부가세","부가세여부","과세구분"]
};
const rules = [
  ["1","의류",["티셔츠","셔츠","블라우스","바지","스커트","원피스","자켓","코트","의류","속옷","양말"]],
  ["2","구두/신발",["신발","구두","운동화","스니커즈","슬리퍼","샌들","부츠"]],
  ["3","가방",["가방","백팩","크로스백","토트백","파우치","캐리어"]],
  ["17","주방용품",["냄비","프라이팬","도마","수저","그릇","컵","텀블러","주방","밀폐용기"]],
  ["18","화장품",["화장품","크림","로션","에센스","세럼","립스틱","샴푸","클렌징"]],
  ["20","농수축산물",["사과","배","과일","채소","농산물","수산물","생선","고기","쌀","계란"]],
  ["21","가공식품",["과자","라면","커피","차","음료","소스","통조림","만두","떡","빵","초콜릿"]],
  ["22","건강기능식품",["건강기능식품","비타민","유산균","오메가","홍삼","영양제"]],
  ["23","어린이제품",["유아","아기","어린이","완구","장난감","유모차","카시트"]],
  ["25","스포츠용품",["스포츠","운동","헬스","골프","축구","야구","테니스","등산","자전거","요가"]],
  ["26","서적",["도서","책","서적","교재","문제집","잡지"]],
  ["40","생활화학제품",["세제","섬유유연제","탈취제","방향제","접착제","코팅제"]]
] as const;
const esmHeaders=["노출사이트","A ID","G ID","상품명","카테고리 템플릿 코드","카테고리 코드","A 노출코드","G 노출코드","판매기간","A 판매가","G 판매가","A 재고","G 재고","옵션 타입","옵션명","옵션 입력값","기본이미지","추가이미지","상품상세설명","배송정보 템플릿 코드","배송방법","출하지 코드","배송정책번호","반품/교환 주소 코드","A 발송정책","G 발송정책","택배사 코드","반품/교환 배송비","상품군 코드","상품고시정보 템플릿코드","원산지 상품타입","원산지 지역타입","원산지 지역코드","복수 원산지여부","청소년구매 불가여부","부가세여부","선물하기상품"];
function clean(v:unknown){return String(v??"").replace(/[\s\n\r_\-()\[\]\/]/g,"").toLowerCase()}
function text(v:unknown){return String(v??"").trim()}
function num(v:unknown){const n=Number(String(v??"").replace(/,/g,"").replace(/[^0-9.-]/g,""));return Number.isFinite(n)?n:0}
function pick(r:Row,names:string[]){for(const n of names){const e=Object.entries(r).find(([k])=>clean(k)===clean(n));if(e&&text(e[1]))return e[1]}return ""}
function parse(sheet:XLSX.WorkSheet){const m=XLSX.utils.sheet_to_json<unknown[]>(sheet,{header:1,defval:"",raw:false});let hi=0,b=-1;for(let i=0;i<Math.min(30,m.length);i++){const s=m[i].filter(v=>Object.values(aliases).flat().some(a=>clean(v)===clean(a))).length;if(s>b){b=s;hi=i}}const h=m[hi].map((v,i)=>text(v)||`열${i+1}`);return m.slice(hi+1).map(r=>Object.fromEntries(h.map((k,i)=>[k,r[i]??""]))).filter(r=>Object.values(r).some(v=>text(v)))}
function products(rows:Row[]):Product[]{return rows.map(raw=>({raw,name:text(pick(raw,aliases.name)),code:text(pick(raw,aliases.code)),price:num(pick(raw,aliases.price)),stock:num(pick(raw,aliases.stock))||99999,image1:text(pick(raw,aliases.image1)),image2:text(pick(raw,aliases.image2)),detail:text(pick(raw,aliases.detail)),optionName:text(pick(raw,aliases.optionName)),optionValue:text(pick(raw,aliases.optionValue)),shipping:num(pick(raw,aliases.shipping)),origin:text(pick(raw,aliases.origin)),vat:text(pick(raw,aliases.vat))})).filter(p=>p.name&&p.price>0)}
function group(p:Product){const hay=text([p.name,...Object.values(p.raw)].join(" ")).toLowerCase();for(const [code,name,ks] of rules)if(ks.some(k=>hay.includes(k)))return {code,name};return {code:"35",name:"기타 재화"}}
function rounded(base:number,margin:number,unit:number){return Math.ceil((base*(1+margin/100))/unit)*unit}
function download(blob:Blob,name:string){const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000)}

export default function Home(){
 const [rows,setRows]=useState<Row[]>([]),[status,setStatus]=useState("상품 엑셀을 업로드해 주세요."),[margin,setMargin]=useState(30),[round,setRound]=useState(100),[aid,setAid]=useState(""),[gid,setGid]=useState(""),[courier,setCourier]=useState<Courier>("CJ대한통운");
 const ps=useMemo(()=>products(rows),[rows]);
 async function read(file?:File){if(!file)return;try{const wb=XLSX.read(await file.arrayBuffer(),{type:"array"});const r=parse(wb.Sheets[wb.SheetNames[0]]);setRows(r);setStatus(`${products(r).length}개 상품을 자동 해석했습니다.`)}catch{setStatus("파일을 읽지 못했습니다.")}}
 function make(chunk:Product[]){const data=[esmHeaders,...chunk.map(p=>{const g=group(p);const raw=p.raw;const val=(...n:string[])=>text(pick(raw,n));return ["옥션/G마켓",aid,gid,p.name,val("카테고리 템플릿 코드"),val("ESM 카테고리코드","카테고리 코드","카테고리코드"),val("A 노출코드","옥션 노출코드"),val("G 노출코드","G마켓 노출코드"),"무제한",rounded(p.price,margin,round),rounded(p.price,margin,round),p.stock,p.stock,p.optionName&&p.optionValue?"단독형":"미사용",p.optionName,p.optionValue,p.image1,p.image2,p.detail,val("배송정보 템플릿 코드"),"일반택배",val("출하지 코드"),val("배송정책번호"),val("반품/교환 주소 코드"),val("A 발송정책","옥션 발송정책"),val("G 발송정책","G마켓 발송정책"),courier==="CJ대한통운"?10013:10007,p.shipping||2500,g.code,val("상품고시정보 템플릿코드","상품고시 템플릿코드"),"해당없음","알수없음",p.origin,"단일원산지","구매가능",p.vat.includes("면세")?"면세상품":"과세상품","가능"]})];const ws=XLSX.utils.aoa_to_sheet(data);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"NEW 일반상품");return wb}
 async function exportEsm(){if(!ps.length)return;if(!aid.trim()&&!gid.trim()){setStatus("옥션 또는 G마켓 판매자 ID를 입력해 주세요.");return}const chunks=[] as Product[][];for(let i=0;i<ps.length;i+=500)chunks.push(ps.slice(i,i+500));const date=new Date().toISOString().slice(0,10);if(chunks.length===1){XLSX.writeFile(make(chunks[0]),`ESM_${ps.length}개_${date}.xlsx`)}else{const zip=new JSZip();chunks.forEach((c,i)=>zip.file(`ESM_${String(i+1).padStart(2,"0")}_${c.length}개.xlsx`,XLSX.write(make(c),{bookType:"xlsx",type:"array"})));download(await zip.generateAsync({type:"blob"}),`ESM_${ps.length}개_분할파일_${date}.zip`)}setStatus(`상품군 자동분류 후 ${ps.length}개 상품 파일을 만들었습니다.`)}
 return <main className="container"><section className="hero"><span className="badge">postsheet02 · ESM 자동변환</span><h1>상품 파일을 ESM 등록파일로 변환</h1><p>상품군은 자동 분류하며 500개씩 자동 분할합니다.</p></section><section className="panel"><div className="field full"><label>상품 엑셀 업로드</label><input type="file" accept=".xlsx,.xls,.csv" onChange={e=>read(e.target.files?.[0])}/></div><div className="field"><label>마진율 (%)</label><input type="number" value={margin} onChange={e=>setMargin(Number(e.target.value))}/></div><div className="field"><label>판매가 올림 단위</label><select value={round} onChange={e=>setRound(Number(e.target.value))}><option value={1}>1원</option><option value={10}>10원</option><option value={100}>100원</option><option value={500}>500원</option><option value={1000}>1,000원</option></select></div><div className="field"><label>옥션 판매자 ID</label><input value={aid} onChange={e=>setAid(e.target.value)}/></div><div className="field"><label>G마켓 판매자 ID</label><input value={gid} onChange={e=>setGid(e.target.value)}/></div><div className="field"><label>택배사</label><select value={courier} onChange={e=>setCourier(e.target.value as Courier)}><option value="CJ대한통운">CJ대한통운 · 10013</option><option value="한진택배">한진택배 · 10007</option></select></div><div className="actions full"><button onClick={exportEsm} disabled={!ps.length}>ESM 파일 다운로드</button></div><div className="status full">{status}</div></section><section className="preview"><div className="previewHead"><h2>자동 분류 미리보기</h2><span>{ps.length}개</span></div><div className="tableWrap"><table><thead><tr><th>상품명</th><th>원래 판매가</th><th>마진 적용가</th><th>상품군</th></tr></thead><tbody>{ps.slice(0,30).map((p,i)=>{const g=group(p);return <tr key={i}><td>{p.name}</td><td>{p.price.toLocaleString()}</td><td>{rounded(p.price,margin,round).toLocaleString()}</td><td>{g.code} · {g.name}</td></tr>})}</tbody></table></div></section></main>
}

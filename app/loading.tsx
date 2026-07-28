export default function Loading() {
  return (
    <main className="container" aria-live="polite">
      <section className="hero">
        <span className="badge">POSTSHEET02</span>
        <h1>배송 엑셀 변환기를 불러오는 중입니다.</h1>
        <p>첫 화면을 먼저 표시한 뒤 엑셀 변환 기능을 준비합니다. 잠시만 기다려 주세요.</p>
        <div className="privacy" style={{ marginTop: 20 }}>
          원본 엑셀과 개인정보는 서버에 저장하지 않습니다.
        </div>
      </section>
    </main>
  );
}

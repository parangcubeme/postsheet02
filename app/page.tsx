"use client";

import dynamic from "next/dynamic";

const ClientApp = dynamic(() => import("./ClientApp"), {
  ssr: false,
  loading: () => (
    <main className="container" aria-live="polite">
      <section className="hero">
        <span className="badge">postsheet02</span>
        <h1>상품 변환기를 준비하고 있습니다.</h1>
        <p>화면을 먼저 표시한 뒤 엑셀 기능을 불러옵니다.</p>
      </section>
    </main>
  ),
});

export default function Page() {
  return <ClientApp />;
}

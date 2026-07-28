import './globals.css';
import './active-tabs.css';
import type { Metadata } from 'next';

export const metadata: Metadata = {
  title: 'postsheet02 상품 엑셀 변환',
  description: '상품 엑셀을 오픈마켓 대량등록 형식으로 변환합니다.',
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return (
    <html lang='ko'>
      <body>{children}</body>
    </html>
  );
}

import { NextResponse } from "next/server";

type Rule = {
  market: string;
  feeRate: number;
  marginRate: number;
  roundUnit: number;
};

const fallbackRules: Rule[] = [
  { market: "스마트스토어", feeRate: 6, marginRate: 30, roundUnit: 100 },
  { market: "옥션", feeRate: 13, marginRate: 30, roundUnit: 100 },
  { market: "G마켓", feeRate: 13, marginRate: 30, roundUnit: 100 },
];

function parseCsvLine(line: string) {
  const result: string[] = [];
  let current = "";
  let quoted = false;

  for (let i = 0; i < line.length; i += 1) {
    const char = line[i];
    if (char === '"') {
      if (quoted && line[i + 1] === '"') {
        current += '"';
        i += 1;
      } else {
        quoted = !quoted;
      }
    } else if (char === "," && !quoted) {
      result.push(current.trim());
      current = "";
    } else {
      current += char;
    }
  }
  result.push(current.trim());
  return result;
}

function parseRules(csv: string): Rule[] {
  const lines = csv.replace(/^\uFEFF/, "").split(/\r?\n/).filter(Boolean);
  if (lines.length < 2) return [];

  const headers = parseCsvLine(lines[0]).map((value) => value.replace(/\s/g, ""));
  const index = {
    market: headers.findIndex((h) => ["마켓", "market", "판매처"].includes(h.toLowerCase())),
    feeRate: headers.findIndex((h) => ["수수료율", "수수료", "feerate"].includes(h.toLowerCase())),
    marginRate: headers.findIndex((h) => ["마진율", "목표마진율", "marginrate"].includes(h.toLowerCase())),
    roundUnit: headers.findIndex((h) => ["올림단위", "반올림단위", "roundunit"].includes(h.toLowerCase())),
  };

  if (Object.values(index).some((value) => value < 0)) return [];

  return lines.slice(1).map(parseCsvLine).map((cells) => ({
    market: cells[index.market] || "",
    feeRate: Number(cells[index.feeRate] || 0),
    marginRate: Number(cells[index.marginRate] || 0),
    roundUnit: Number(cells[index.roundUnit] || 100),
  })).filter((rule) => rule.market && Number.isFinite(rule.feeRate) && Number.isFinite(rule.marginRate));
}

export async function GET() {
  const csvUrl = process.env.GOOGLE_SHEET_CSV_URL;
  if (!csvUrl) {
    return NextResponse.json({ source: "fallback", rules: fallbackRules });
  }

  try {
    const response = await fetch(csvUrl, { cache: "no-store" });
    if (!response.ok) throw new Error(`Google Sheet response: ${response.status}`);
    const rules = parseRules(await response.text());
    return NextResponse.json({ source: rules.length ? "google-sheet" : "fallback", rules: rules.length ? rules : fallbackRules });
  } catch {
    return NextResponse.json({ source: "fallback", rules: fallbackRules });
  }
}

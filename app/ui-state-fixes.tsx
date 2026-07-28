"use client";

import { useEffect } from "react";

export default function UiStateFixes() {
  useEffect(() => {
    const sync = () => {
      const buttons = Array.from(document.querySelectorAll("button"));
      const stepButtons = buttons.filter(button => /^\d+\./.test(button.textContent?.trim() || ""));
      const selectedStep = stepButtons.find(button => getComputedStyle(button).fontWeight === "700");
      stepButtons.forEach(button => button.setAttribute("aria-pressed", String(button === selectedStep)));

      const smart = buttons.find(button => button.textContent?.trim() === "스마트스토어");
      const esm = buttons.find(button => button.textContent?.includes("ESM · 옥션/G마켓"));
      const esmOpen = stepButtons.length > 0;
      smart?.setAttribute("aria-pressed", String(!esmOpen));
      esm?.setAttribute("aria-pressed", String(esmOpen));

      const labels = Array.from(document.querySelectorAll("label"));
      const auction = labels.find(label => label.textContent?.includes("옥션 판매자 ID"));
      const gmarket = labels.find(label => label.textContent?.includes("G마켓 판매자 ID"));
      const idArea = auction?.parentElement;
      if (idArea && gmarket?.parentElement === idArea) {
        const reviewOpen = selectedStep?.textContent?.includes("최종검사") ?? false;
        idArea.style.display = reviewOpen ? "grid" : "none";
      }
    };

    sync();
    const observer = new MutationObserver(sync);
    observer.observe(document.body, { childList: true, subtree: true, attributes: true });
    document.addEventListener("click", sync, true);
    return () => {
      observer.disconnect();
      document.removeEventListener("click", sync, true);
    };
  }, []);

  return null;
}

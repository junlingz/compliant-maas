import { chromium } from "/Users/a1/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/playwright/index.mjs";
import { mkdir, readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";

const base = process.env.PROTOTYPE_URL || "http://127.0.0.1:4173";
const out = new URL("../artifacts/", import.meta.url);
await mkdir(out, { recursive: true });
const screenshotPath = (name) => fileURLToPath(new URL(name, out));
const browser = await chromium.launch({
  headless: true,
  executablePath: "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
});
const page = await browser.newPage({ viewport: { width: 1440, height: 1000 } });
const failures = [];
const consoleErrors = [];
page.on("console", (message) => {
  if (message.type() === "error") consoleErrors.push(message.text());
});
page.on("pageerror", (error) => consoleErrors.push(error.message));
const check = (condition, message) => {
  if (!condition) failures.push(message);
};

await page.goto(base, { waitUntil: "networkidle" });
await page.screenshot({ path: screenshotPath("overview-desktop.png"), fullPage: true });

const moduleIds = [
  "autoregressive",
  "seq2seq",
  "text2image",
  "distributed",
  "finetune",
  "rl",
  "evaluation",
  "inference",
];
let sectionCount = 0;
let requirementCards = 0;
for (const id of moduleIds) {
  await page.locator(`[data-page="${id}"]`).first().click();
  const count = await page.locator("[data-section]").count();
  check(count > 0, `${id}: no section tabs`);
  for (let index = 0; index < count; index += 1) {
    await page.locator("[data-section]").nth(index).click();
    sectionCount += 1;
    const cards = await page.locator(".req").count();
    requirementCards += cards;
    check(cards > 0, `${id} section ${index}: no requirement cards`);
    check((await page.locator(".work-card h2").count()) > 0, `${id} section ${index}: no workbench title`);

    const targetResult = await page.evaluate(() => {
      const reqButton = document.querySelector(".req [data-requirement]");
      reqButton?.click();
      const target = document.activeElement;
      return {
        exists: Boolean(target?.closest?.(".work-card")),
        id: target?.id || "",
        controlId: target?.dataset.controlId || "",
        focused: target === document.activeElement,
        rightCardHasReqId: Boolean(document.querySelector(".audit-card [data-req-id]")),
      };
    });
    check(targetResult.exists && targetResult.controlId.startsWith("control-"), `${id} section ${index}: requirement not mapped to control-*`);
    check(targetResult.focused, `${id} section ${index}: requirement locator did not focus business control`);
    check(!targetResult.rightCardHasReqId, `${id} section ${index}: audit card still masquerades as requirement target`);

    const unnamed = await page.evaluate(() =>
      [...document.querySelectorAll(".work-card input,.work-card select,.work-card textarea")]
        .filter((el) => el.offsetParent !== null)
        .filter((el) => {
          const label = el.id && document.querySelector(`label[for="${CSS.escape(el.id)}"]`);
          return !el.getAttribute("aria-label") && !el.getAttribute("aria-labelledby") && !label;
        })
        .map((el) => `${el.tagName.toLowerCase()}#${el.id || "(no-id)"}`)
    );
    check(unnamed.length === 0, `${id} section ${index}: unnamed controls ${unnamed.join(", ")}`);
  }
}
check(sectionCount === 39, `expected 39 sections, got ${sectionCount}`);
check(requirementCards === 143, `expected 143 feature cards, got ${requirementCards}`);

// Autoregressive configuration validation, export and real detail drawer.
await page.locator('[data-page="autoregressive"]').first().click();
await page.locator('[data-section="0"]').click();
await page.locator("#architectureSelect").selectOption({ label: "MoE 稀疏专家" });
check(await page.locator("#moePanel").isVisible(), "MoE parameters did not appear");
check((await page.locator("#parameterSummary").textContent()).includes("Experts"), "MoE summary did not update");
await page.locator("#learningRate").fill("0.1");
await page.locator('[data-action="startTask"]').first().click();
check((await page.locator("#modalBackdrop.open").count()) === 0, "invalid task submission opened success modal");
await page.locator("#learningRate").fill("0.00002");
await page.locator("#batchSize").selectOption("32");
const configDownloadPromise = page.waitForEvent("download");
await page.locator('[data-action="exportConfig"]').first().click();
const configDownload = await configDownloadPromise;
const exportedConfig = JSON.parse(await readFile(await configDownload.path(), "utf8"));
check(exportedConfig.lr === "0.00002" && exportedConfig.batch === "32", "configuration export ignored current values");
await page.locator('[data-section="1"]').click();
await page.locator('[data-action="preview"]').first().click();
check((await page.locator("#drawerBackdrop.open").count()) === 1, "business detail drawer did not open");
check(await page.locator("#drawer").evaluate((el) => el.contains(document.activeElement)), "drawer did not receive focus");
await page.keyboard.press("Escape");
check((await page.locator("#drawerBackdrop.open").count()) === 0, "Escape did not close drawer");

// Seq2Seq architecture details are visible and precise.
await page.locator('[data-page="seq2seq"]').first().click();
await page.locator('[data-section="1"]').click();
const seqDetails = await page.locator("#seq2seqArchitectureDetail").innerText();
check(seqDetails.includes("406M") && seqDetails.includes("16"), "Seq2Seq architecture lacks layers/params/heads");

// Text-image parameter preset changes actual fields.
await page.locator('[data-page="text2image"]').first().click();
await page.locator('[data-section="2"]').click();
const presetBefore = await page.locator(".form-grid input").first().inputValue();
await page.locator('[data-action="applyPreset"]').nth(1).click();
check((await page.locator(".form-grid input").first().inputValue()) !== presetBefore, "hyperparameter preset did not update inputs");

// Distributed global performance, sortable nodes and dynamic aggregation.
await page.locator('[data-page="distributed"]').first().click();
await page.locator('[data-section="0"]').click();
check((await page.locator(".chart").count()) > 0, "distributed core lacks global performance chart");
check((await page.locator("#nodePerformance tr").count()) >= 2, "distributed core lacks per-node table");
const firstNodeBefore = await page.locator("#nodePerformance tr").first().innerText();
await page.locator('[data-action="sortNodes"]').click();
check((await page.locator("#nodePerformance tr").first().innerText()) !== firstNodeBefore, "node sorting did not change table");
await page.locator("#aggregationSlider").evaluate((el) => {
  el.value = "8";
  el.dispatchEvent(new Event("input", { bubbles: true }));
});
check((await page.locator("#aggregationSlider + .range-value").textContent()).includes("8"), "aggregation control did not update");
await page.locator('[data-section="1"]').click();
check((await page.locator("[data-parallel]").count()) === 4, "parallel workbench missing four paradigms");
await page.locator('[data-action="validateParallel"]').click();
check((await page.locator("#parallelPreview").textContent()).includes("passed"), "parallel validation did not update preview");

// Fine-tuning has chart, terminal and final report with performance/time/resources.
await page.locator('[data-page="finetune"]').first().click();
await page.locator('[data-section="0"]').click();
check((await page.locator(".chart").count()) > 0 && (await page.locator("#liveLog").count()) > 0, "fine-tuning core lacks chart or terminal");
await page.locator('[data-action="previewFinetuneReport"]').click();
check(await page.locator("#finetuneReport").isVisible(), "fine-tuning report did not reveal");
const fineReport = await page.locator("#finetuneReport").innerText();
check(/F1|准确率/.test(fineReport) && /时长|GPU|显存|能耗/.test(fineReport), "fine-tuning report lacks metrics/time/resources");
await page.locator('[data-section="2"]').click();
check((await page.locator("text=GPU-0").count()) > 0, "GPU schedule page lacks per-GPU metrics");

// RL history and zoom change actual chart state and final metrics.
await page.locator('[data-page="rl"]').first().click();
await page.locator('[data-section="1"]').click();
const rlPathBefore = await page.locator(".chart path[stroke]").first().getAttribute("d");
await page.locator("#monitorTask").selectOption({ index: 1 });
check((await page.locator(".chart path[stroke]").first().getAttribute("d")) !== rlPathBefore, "RL history did not change chart");
check((await page.locator("#liveStep").textContent()).includes("最终"), "RL history did not expose final metrics");
const viewBoxBefore = await page.locator(".chart svg").getAttribute("viewBox");
await page.locator('[data-action="zoomChart"]').click();
check((await page.locator(".chart svg").getAttribute("viewBox")) !== viewBoxBefore, "RL zoom did not change viewBox");

// Evaluation type linkage and report-specific visualization.
await page.locator('[data-page="evaluation"]').first().click();
await page.locator('[data-section="0"]').click();
await page.locator('[data-model-type="多模态模型"]').click();
check((await page.locator("#evalTask").textContent()).includes("视觉问答"), "evaluation model type did not update tasks");
await page.locator('[data-section="2"]').click();
check((await page.getByText("混淆矩阵").count()) > 0, "evaluation report lacks confusion matrix");

// Inference asset CRUD, logs, analytics and service panels.
await page.locator('[data-page="inference"]').first().click();
await page.locator('[data-section="0"]').click();
await page.locator("#modelSearch").fill("不存在模型XYZ");
await page.locator('[data-action="filterModels"]').click();
check(await page.locator("#assetEmpty").isVisible(), "model filter did not show empty state");
await page.locator('[data-action="editCategories"]').click();
check(await page.locator("#assetAdminPanel").isVisible(), "asset admin panel did not open");
await page.locator("#newAssetName").fill("QA-New-Model");
const assetCountBefore = await page.locator(".asset-card").count();
await page.locator('[data-action="saveAssetMetadata"]').click();
check((await page.locator(".asset-card").count()) === assetCountBefore + 1, "asset CRUD did not add model");
await page.locator("#experienceLogSearch").fill("researcher");
await page.locator('[data-action="experienceLogs"]').click();
check((await page.locator("#experienceLogRows tr:visible").count()) === 1, "experience log filter did not change rows");
await page.locator('[data-action="usageAnalytics"]').click();
check(await page.locator("#assetAnalytics").isVisible(), "usage analytics did not open in page");
await page.locator('[data-section="2"]').click();
await page.locator('[data-action="auditLogs"]').last().click();
check(await page.locator("#serviceDetailPanel").isVisible(), "service audit panel did not open");
check((await page.locator("#serviceDetailRows tr").count()) === 2, "service audit panel lacks rows");
await page.locator('[data-action="orchestrationHistory"]').last().click();
check((await page.locator("#serviceDetailTitle").textContent()).includes("编排"), "orchestration history did not replace panel");
check((await page.locator("#serviceDetailRows").innerText()).includes("ORCH-"), "orchestration history lacks instances");

// Documents: empty state and valid PDF.
await page.locator('[data-page="docs"]').first().click();
await page.locator("#docSearch").fill("绝对不存在的搜索词XYZ");
await page.locator('[data-action="searchDocs"]').click();
check(await page.locator("#docEmpty").isVisible(), "document search did not show empty state");
const pdfDownloadPromise = page.waitForEvent("download");
await page.locator('[data-action="downloadWhitepaper"]').click();
const pdfDownload = await pdfDownloadPromise;
const pdfBytes = await readFile(await pdfDownload.path());
check(pdfDownload.suggestedFilename().endsWith(".pdf") && pdfBytes.subarray(0, 4).equals(Buffer.from("%PDF")), "whitepaper is not a valid PDF");

// Modal focus, Escape and requirement-to-business-control locator.
await page.locator('[data-page="autoregressive"]').first().click();
await page.locator('[data-section="0"]').click();
await page.locator("#learningRate").fill("0.00002");
await page.locator('[data-action="startTask"]').first().click();
await page.waitForTimeout(30);
check((await page.locator("#modalBackdrop.open").count()) === 1, "valid task submission modal did not open");
check(await page.locator("#modal").evaluate((el) => el.contains(document.activeElement)), "modal did not receive focus");
await page.keyboard.press("Escape");
check((await page.locator("#modalBackdrop.open").count()) === 0, "Escape did not close modal");
await page.locator(".req [data-requirement]").first().click();
const located = await page.evaluate(() => {
  const el = document.activeElement;
  return { controlId: el?.dataset.controlId || "", reqId: el?.dataset.reqId || "", focused: el?.classList.contains("req-focus") };
});
check(located.controlId.startsWith("control-") && located.reqId && located.focused, "requirement did not locate an actual business control");

await page.locator('[data-page="coverage"]').first().click();
await page.locator("#coverageSearch").fill("模型版本");
await page.locator('[data-action="searchCoverage"]').click();
check((await page.locator("#coverageBody tr").count()) > 0, "coverage search returned no rows");
check((await page.locator("#coverageBody").innerText()).includes("data-control-id"), "coverage matrix still points to audit cards");
await page.screenshot({ path: screenshotPath("coverage-desktop.png"), fullPage: true });

// Mobile navigation: scrim, focus trap, aria-expanded reset and no overflow.
await page.setViewportSize({ width: 390, height: 844 });
await page.goto(`${base}/#overview`, { waitUntil: "networkidle" });
const menuButton = page.locator('[data-action="menu"]');
await menuButton.click();
check((await page.locator("#sidebar.open").count()) === 1, "mobile menu did not open");
check((await page.locator("#mobileScrim.open").count()) === 1, "mobile scrim did not open");
check((await menuButton.getAttribute("aria-expanded")) === "true", "mobile menu aria-expanded not true when open");
await page.locator("#nav button").last().focus();
await page.keyboard.press("Tab");
check(await page.locator("#sidebar").evaluate((el) => el.contains(document.activeElement)), "mobile nav focus escaped to background");
await page.getByRole("button", { name: "关闭导航" }).click();
check((await menuButton.getAttribute("aria-expanded")) === "false", "mobile menu aria-expanded not reset");
await page.screenshot({ path: screenshotPath("overview-mobile.png"), fullPage: true });

await page.setViewportSize({ width: 320, height: 700 });
await page.goto(`${base}/#evaluation`, { waitUntil: "networkidle" });
const widths = await page.evaluate(() => ({ scroll: document.documentElement.scrollWidth, inner: innerWidth }));
check(widths.scroll === widths.inner, `320px viewport overflow: ${JSON.stringify(widths)}`);

// Standalone file:// execution.
const filePage = await browser.newPage({ viewport: { width: 1200, height: 800 } });
const fileErrors = [];
filePage.on("pageerror", (error) => fileErrors.push(error.message));
await filePage.goto(new URL("../index.html", import.meta.url).href, { waitUntil: "load" });
check((await filePage.locator("#nav button").count()) >= 12, "file:// standalone navigation failed");
check(fileErrors.length === 0, `file:// errors: ${fileErrors.join("; ")}`);
await filePage.close();

await browser.close();
const report = {
  moduleCount: moduleIds.length,
  sectionCount,
  requirementCardsVisited: requirementCards,
  consoleErrors,
  failures,
  passed: failures.length === 0 && consoleErrors.length === 0,
};
console.log(JSON.stringify(report, null, 2));
if (!report.passed) process.exitCode = 1;

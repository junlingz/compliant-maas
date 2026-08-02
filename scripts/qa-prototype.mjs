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
let navigationSequence = 0;
const freshUrl = (hash = "") => {
  const url = new URL(base);
  url.searchParams.set("qa", `${Date.now()}-${navigationSequence += 1}`);
  url.hash = hash;
  return url.href;
};
const failures = [];
const consoleErrors = [];
page.on("console", (message) => {
  if (message.type() === "error") consoleErrors.push(message.text());
});
page.on("pageerror", (error) => consoleErrors.push(error.message));
const check = (condition, message) => {
  if (!condition) failures.push(message);
};

await page.goto(freshUrl(), { waitUntil: "networkidle" });
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
let productionScreensVisited = 0;
const internalCopy = /标书|本原型|条款|合规覆盖|需求映射|能力核验|覆盖完整|已验证覆盖|可核验点/;
check(!internalCopy.test(await page.locator("body").innerText()), "overview exposes internal bid/review copy");
check((await page.locator(".scope,.audit-card,.hero").count()) === 0, "overview still renders prototype explanation surfaces");
check((await page.locator('[data-page="coverage"]').count()) === 0, "overview still exposes an internal coverage entry");
check((await page.locator('.workspace-switcher,[data-action="workspace"]').count()) === 0, "removed pretraining production workspace is visible");
check(!(await page.locator("body").innerText()).includes("预训练生产空间"), "removed pretraining production workspace copy is visible");
for (const id of moduleIds) {
  await page.locator(`[data-page="${id}"]`).first().click();
  const count = await page.locator("[data-section]").count();
  check(count > 0, `${id}: no section tabs`);
  for (let index = 0; index < count; index += 1) {
    await page.locator("[data-section]").nth(index).click();
    sectionCount += 1;
    productionScreensVisited += 1;
    check((await page.locator(".work-card h2").count()) > 0, `${id} section ${index}: no workbench title`);
    check((await page.locator(".scope,.audit-card,.hero,.req").count()) === 0, `${id} section ${index}: internal explanation UI is visible`);
    check(!internalCopy.test(await page.locator("body").innerText()), `${id} section ${index}: internal bid/review copy is visible`);

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
check(productionScreensVisited === 39, `expected 39 production screens, got ${productionScreensVisited}`);

// Remote algorithm-service board must survive the superscale merge unchanged and remain routable.
const algorithmModules = [
  ["graphcompute", "图计算算法工具集", 3],
  ["promptlearning", "提示学习算法", 5],
  ["reverseprompt", "反向提示算法", 3],
  ["fewshot", "小样本学习算法", 4],
];
for (const [id, title, expectedSections] of algorithmModules) {
  check((await page.locator(`[data-page="${id}"]`).count()) > 0, `${title}: navigation was overwritten during merge`);
  await page.locator(`[data-page="${id}"]`).first().click();
  check((await page.locator(".page-toolbar h1").textContent()) === title, `${title}: route renders the wrong page`);
  check((await page.locator(`[data-module="${id}"][data-section]`).count()) === expectedSections, `${title}: expected ${expectedSections} sections`);
  check((await page.locator(".work-card h2").count()) > 0, `${title}: workbench content is missing`);
}

// Autoregressive configuration validation, export and real detail drawer.
await page.locator('[data-page="autoregressive"]').first().click();
await page.locator('[data-section="0"]').click();
await page.screenshot({ path: screenshotPath("workbench-desktop.png"), fullPage: true });
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
const modelChooserPromise = page.waitForEvent("filechooser");
const importedModelCountBefore = await page.locator(".work-card .model-grid .model-card").count();
await page.locator('[data-action="upload"][data-context="model"]').click();
const modelChooser = await modelChooserPromise;
check(await page.locator("#hiddenFile").getAttribute("multiple") !== null, "model import does not allow weight + config selection");
await modelChooser.setFiles([
  { name: "weights.safetensors", mimeType: "application/octet-stream", buffer: Buffer.from("weights") },
  { name: "config.json", mimeType: "application/json", buffer: Buffer.from('{"model_type":"decoder"}') },
]);
check((await page.locator("#modal").innerText()).includes("权重") && (await page.locator("#modal").innerText()).includes("配置"), "model import did not pair weight and config");
await page.locator('[data-action="confirmModal"]').click();
check((await page.locator(".work-card .model-grid .model-card").count()) === importedModelCountBefore + 1, "validated model import did not create a model/version card");
check((await page.locator("#selectedModelState").innerText()).includes("weights · v1.0"), "imported model was not selected as the current version");
check((await page.locator("#modalBackdrop.open").count()) === 0 && (await page.locator("#modal").textContent()) === "", "model import confirmation left modal residue");

// Upload parsing rejects malformed JSONL instead of validating by extension alone.
await page.locator('[data-page="autoregressive"]').first().click();
await page.locator('[data-section="0"]').click();
const invalidJsonlChooserPromise = page.waitForEvent("filechooser");
await page.locator('[data-action="upload"][data-context="autoregressive-dataset"]').click();
const invalidJsonlChooser = await invalidJsonlChooserPromise;
await invalidJsonlChooser.setFiles({
  name: "broken.jsonl",
  mimeType: "application/x-ndjson",
  buffer: Buffer.from('{"text":"valid"}\n{"text":'),
});
await page.waitForTimeout(30);
check((await page.locator("#modalTitle").textContent()) === "JSONL 解析失败", "malformed multiline JSONL was accepted");
await page.locator('[data-action="closeModal"]').first().click();

// Seq2Seq architecture details are visible and precise.
await page.locator('[data-page="seq2seq"]').first().click();
await page.locator('[data-section="1"]').click();
const seqDetails = await page.locator("#seq2seqArchitectureDetail").innerText();
check(seqDetails.includes("406M") && seqDetails.includes("16"), "Seq2Seq architecture lacks layers/params/heads");
await page.locator('[data-section="0"]').click();
check((await page.locator("#seq2seqMonitorPanel").count()) === 1, "Seq2Seq submission lacks a dedicated dynamic monitor panel");

// Source/target corpora are parsed by line and mismatches block submission.
await page.locator('[data-section="0"]').click();
const sourceChooserPromise = page.waitForEvent("filechooser");
await page.locator('[data-action="upload"][data-context="seq2seq-src"]').click();
const sourceChooser = await sourceChooserPromise;
await sourceChooser.setFiles({
  name: "src.txt",
  mimeType: "text/plain",
  buffer: Buffer.from("第一行\n第二行"),
});
await page.waitForTimeout(30);
await page.locator('[data-action="confirmModal"]').click();
const targetChooserPromise = page.waitForEvent("filechooser");
await page.locator('[data-action="upload"][data-context="seq2seq-tgt"]').click();
const targetChooser = await targetChooserPromise;
await targetChooser.setFiles({
  name: "tgt.txt",
  mimeType: "text/plain",
  buffer: Buffer.from("only one line"),
});
await page.waitForTimeout(30);
check((await page.locator("#modalTitle").textContent()) === "语料对齐失败", "unequal Seq2Seq line counts were accepted");
check((await page.locator("#alignmentStatus").innerText()).includes("2 / 1") && (await page.locator("#alignmentStatus").innerText()).includes("不一致"), "Seq2Seq alignment status did not expose actual line counts");
await page.locator('[data-action="closeModal"]').first().click();
await page.locator('[data-action="startTask"]').first().click();
check((await page.locator("#modalBackdrop.open").count()) === 0, "Seq2Seq submission was not blocked after alignment failure");

// Text-image parameter preset changes actual fields.
await page.locator('[data-page="text2image"]').first().click();
await page.locator('[data-section="2"]').click();
const presetBefore = await page.locator(".form-grid input").first().inputValue();
await page.locator('[data-action="applyPreset"]').nth(1).click();
check((await page.locator(".form-grid input").first().inputValue()) !== presetBefore, "hyperparameter preset did not update inputs");
await page.locator('[data-section="0"]').click();
const invalidImageChooserPromise = page.waitForEvent("filechooser");
await page.locator('[data-action="upload"][data-context="image-pairs"]').click();
const invalidImageChooser = await invalidImageChooserPromise;
await invalidImageChooser.setFiles({
  name: "damaged.jpg",
  mimeType: "image/jpeg",
  buffer: Buffer.from("not-a-decodable-image"),
});
await page.waitForTimeout(100);
check((await page.locator("#modalTitle").textContent()) === "图像校验失败", "damaged image was accepted using extension/size only");
await page.locator('[data-action="closeModal"]').first().click();

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
await page.locator('[data-section="2"]').click();
check((await page.getByText("100B 混合并行训练").count()) === 1 && (await page.getByText("自定义通信插件").count()) === 1, "distributed scenario examples are missing");
await page.locator('[data-action="runDistributedExample"]').first().click();
check((await page.locator("#distributedExampleOutput").innerText()).includes("91.8%"), "distributed example did not produce expected output");

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
await page.locator('[data-section="1"]').click();
check((await page.locator('[data-failed-finetune]').count()) === 1, "fine-tuning queue lacks a failed task");
await page.locator('[data-queue-action="resubmit"]').click();
check((await page.locator('[data-failed-finetune]').innerText()).includes("按原配置重新排队"), "failed fine-tuning task was not resubmitted with original config");
await page.locator('[data-section="3"]').click();
check((await page.getByText("自定义优化器扩展模板").count()) > 0, "extension library lacks optimizer template");
check((await page.getByText("扩展开发工具").count()) > 0 && (await page.getByText("教程与完整案例").count()) > 0, "extension tooling or tutorials are missing");

// RL history and zoom change actual chart state and final metrics.
await page.locator('[data-page="rl"]').first().click();
await page.locator('[data-section="0"]').click();
check((await page.locator("#rlSwaggerPanel").count()) === 1, "RL API lacks Swagger-style parameter browser");
await page.locator('[data-action="tryRlApi"]').click();
check((await page.locator("#rlTryResponse").innerText()).includes("queued"), "RL Try action did not return a response");
await page.locator('[data-section="1"]').click();
check((await page.getByText("强化学习历史任务").count()) === 1 && (await page.getByText("09:42:18").count()) > 0, "RL monitor lacks algorithm/duration/status history");
const rlPathBefore = await page.locator(".chart path[stroke]").first().getAttribute("d");
await page.locator("#monitorTask").selectOption({ index: 1 });
check((await page.locator(".chart path[stroke]").first().getAttribute("d")) !== rlPathBefore, "RL history did not change chart");
check((await page.locator("#liveStep").textContent()).includes("最终"), "RL history did not expose final metrics");
const viewBoxBefore = await page.locator(".chart svg").getAttribute("viewBox");
await page.locator('[data-action="zoomChart"]').click();
check((await page.locator(".chart svg").getAttribute("viewBox")) !== viewBoxBefore, "RL zoom did not change viewBox");
await page.locator('[data-section="2"]').click();
await page.locator('[data-action="rlStart"]').click();
check((await page.locator("#rlLifecycleStatus").textContent()) === "运行中", "RL lifecycle start API did not update status");
await page.locator('[data-action="rlTerminate"]').click();
check((await page.locator("#rlLifecycleResponse").innerText()).includes("terminate"), "RL lifecycle terminate API did not return response");

// Evaluation type linkage and report-specific visualization.
await page.locator('[data-page="evaluation"]').first().click();
await page.locator('[data-section="0"]').click();
await page.locator('[data-model-type="多模态模型"]').click();
check((await page.locator("#evalTask").textContent()).includes("视觉问答"), "evaluation model type did not update tasks");
await page.locator('[data-section="2"]').click();
check((await page.getByText("混淆矩阵").count()) > 0, "evaluation report lacks confusion matrix");
check((await page.locator(".radar-chart svg").count()) === 1, "evaluation report does not render a real radar chart");
await page.screenshot({ path: screenshotPath("evaluation-report-desktop.png"), fullPage: true });
await page.locator('[data-section="1"]').click();
check((await page.locator("#evalTaskModelType").count()) === 1 && (await page.locator("#evalTaskOperator").count()) === 1 && (await page.locator("#evalTaskDateRange").count()) === 1, "evaluation task list lacks model/operator/date filters");

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
for (const [id, expectedTitle] of [["autoregressive", "自回归预训练框架技术白皮书"], ["seq2seq", "序列到序列预训练框架技术白皮书"], ["text2image", "文-图生成训练框架技术白皮书"], ["distributed", "拓扑与资源感知分布式训练技术白皮书"]]) {
  await page.locator(`[data-page="${id}"]`).first().click();
  await page.locator(`[data-module="${id}"][data-section]`).filter({ hasText: "技术文档" }).click();
  check((await page.locator('[data-acceptance="framework-whitepaper"]').innerText()).includes(expectedTitle), `${id}: dedicated whitepaper is missing`);
}
await page.locator('[data-page="finetune"]').first().click();
await page.locator('[data-module="finetune"][data-section]').filter({ hasText: "技术文档" }).click();
check((await page.locator('[data-doc-section="finetuneGuide"]').count()) === 1 && (await page.locator('[data-doc-section="finetuneEvaluation"]').count()) === 1, "fine-tuning docs lack guide/evaluation chapters");
await page.locator('[data-doc-section="finetuneEvaluation"]').click();
check((await page.locator("#docReader").innerText()).includes("PPL") && (await page.locator("#docReader").innerText()).includes("Recall"), "fine-tuning evaluation chapter lacks required metrics/formulas");
await page.locator('[data-page="evaluation"]').first().click();
await page.locator('[data-module="evaluation"][data-section]').filter({ hasText: "技术文档" }).click();
check((await page.locator('[data-doc-section="evaluationMethods"]').count()) === 1, "evaluation docs lack dedicated model evaluation methods chapter");
await page.locator('[data-page="docs"]').first().click();
await page.locator("#docSearch").fill("绝对不存在的搜索词XYZ");
await page.locator('[data-action="searchDocs"]').click();
check(await page.locator("#docEmpty").isVisible(), "document search did not show empty state");
const pdfDownloadPromise = page.waitForEvent("download");
await page.locator('[data-action="downloadWhitepaper"]').click();
const pdfDownload = await pdfDownloadPromise;
const pdfBytes = await readFile(await pdfDownload.path());
check(pdfDownload.suggestedFilename().endsWith(".pdf") && pdfBytes.subarray(0, 4).equals(Buffer.from("%PDF")), "whitepaper is not a valid PDF");

// Modal focus and Escape.
await page.locator('[data-page="autoregressive"]').first().click();
await page.locator('[data-section="0"]').click();
await page.locator("#learningRate").fill("0.00002");
await page.locator('[data-action="startTask"]').first().click();
await page.waitForTimeout(30);
check((await page.locator("#modalBackdrop.open").count()) === 1, "valid task submission modal did not open");
check(await page.locator("#modal").evaluate((el) => el.contains(document.activeElement)), "modal did not receive focus");
await page.keyboard.press("Escape");
check((await page.locator("#modalBackdrop.open").count()) === 0, "Escape did not close modal");
check((await page.locator("#modal").textContent()) === "", "closed modal retained stale content");

// Release-critical requirements: editable training controls and new-task monitor identity.
check((await page.locator("#optimizer").count()) === 1, "autoregressive optimizer is not editable");
check((await page.locator("#lrSchedule").count()) === 1, "autoregressive LR schedule is missing");
check((await page.locator("#weightDecay").count()) === 1 && (await page.locator("#gradientClip").count()) === 1, "autoregressive stability parameters missing");
check((await page.locator("#checkpointInterval").count()) === 1 && (await page.locator("#checkpointRetention").count()) === 1, "autoregressive checkpoint controls missing");
await page.locator('[data-action="startTask"]').first().click();
const createdTaskText = await page.locator("#modal").innerText();
const createdTaskId = createdTaskText.match(/PT-\d{8}-\d{3}/)?.[0];
await page.locator('[data-action="confirmModal"]').click();
check(Boolean(createdTaskId), "new task modal did not expose task ID");
check((await page.locator("#monitorTaskId").textContent()) === createdTaskId, "monitor did not open the newly created task");
check((await page.locator("#modalBackdrop.open").count()) === 0, "confirm left modal overlay visible");

// Seq2Seq pretraining/downstream and fine-tuning parameters are operational.
await page.locator('[data-page="seq2seq"]').first().click();
await page.locator('[data-section="0"]').click();
check((await page.locator("#seqLearningRate").count()) === 1 && (await page.locator("#seqBatch").count()) === 1 && (await page.locator("#seqEpochs").count()) === 1, "Seq2Seq pretraining parameters missing");
await page.locator('[data-action="applySeqPreset"]').click();
check((await page.locator("#seqEpochs").inputValue()) === "12", "Seq2Seq preset did not apply");
await page.locator('[data-section="4"]').click();
check((await page.locator("#downstreamCheckpoint option").count()) >= 3, "downstream checkpoint selection missing");
check((await page.locator("#downstreamOptimizer").count()) === 1 && (await page.locator("#downstreamBatchEpochs").count()) === 1, "downstream hyperparameters missing");
await page.locator('[data-downstream-type="文本摘要"]').check();
check((await page.locator("#downstreamDataset").inputValue()) === "summary_train.jsonl", "downstream task switch did not link the dataset");
check((await page.locator("#downstreamPreset").inputValue()) === "长文本摘要", "downstream task switch did not link the parameter preset");
await page.locator("#testInput").fill("请概括：第一份输入强调训练吞吐。");
await page.locator('[data-action="generateTest"]').click();
const downstreamOutputBefore = await page.locator("#testOutput").innerText();
await page.locator("#testInput").fill("请概括：第二份输入强调部署延迟和稳定性。");
await page.locator('[data-action="generateTest"]').click();
check((await page.locator("#testOutput").innerText()) !== downstreamOutputBefore, "downstream online test returned the same output for different inputs");
await page.locator('[data-page="finetune"]').first().click();
await page.locator('[data-section="0"]').click();
check((await page.locator('[name="finetuneModel"]').count()) === 3, "fine-tuning model inspector cards missing");
check((await page.locator("#finetuneEpochs").count()) === 1 && (await page.locator("#finetuneLearningRate").count()) === 1 && (await page.locator("#finetuneBatch").count()) === 1, "fine-tuning core parameters missing");
await page.locator('[data-action="applyFinetuneRecommendation"]').click();
check((await page.locator("#finetuneEpochs").inputValue()) === "4", "fine-tuning recommendation did not update parameters");

// Task detail includes config, logs, timeline and exact log export.
await page.locator('[data-page="tasks"]').first().click();
await page.locator('[data-task-action="detail"]').first().click();
check(await page.locator("#taskDetailPanel").isVisible(), "task detail panel did not open");
check((await page.locator("#taskConfigJson").innerText()).includes("framework"), "task detail lacks submitted configuration");
check((await page.locator("#taskTimeline .timeline-item").count()) >= 3, "task detail lacks status timeline");
await page.locator("#taskLogLevel").selectOption("WARNING");
await page.locator('[data-action="filterTaskLogs"]').click();
check((await page.locator("#taskDetailLogs").innerText()).includes("WARNING"), "task log level filter failed");
const taskLogDownloadPromise = page.waitForEvent("download");
await page.locator('[data-action="exportTaskLogs"]').first().click();
const taskLogDownload = await taskLogDownloadPromise;
check(taskLogDownload.suggestedFilename().endsWith("-logs.txt"), "task log export downloaded the wrong artifact");
await page.locator("#taskSearch").fill("definitely-not-a-task");
await page.locator('[data-action="filterTasks"]').click();
check(await page.locator("#taskEmpty").isVisible(), "task filters lack an explicit empty state");
await page.locator('[data-action="clearTaskFilters"]').click();

// Evaluation linkage, completed-task comparison and custom dataset metadata.
await page.locator('[data-page="evaluation"]').first().click();
await page.locator('[data-section="0"]').click();
await page.locator('[data-model-type="多模态模型"]').click();
check((await page.locator("#evalDataset").textContent()).includes("MMBench"), "multimodal type did not update evaluation datasets");
await page.locator('[data-section="3"]').click();
await page.locator('[data-compare-model="GLM"]').check();
await page.locator('[data-action="runModelComparison"]').click();
check((await page.locator("#compareHead th").count()) === 5, "model comparison did not generate dynamic columns");
await page.locator('[data-section="4"]').click();
await page.locator('[data-action="openDatasetUpload"]').click();
check(await page.locator("#datasetUploadPanel").isVisible(), "custom dataset metadata editor did not open");
check((await page.locator("#datasetName").count()) === 1 && (await page.locator("#datasetInputMap").count()) === 1, "custom dataset metadata/schema controls missing");
const invalidDatasetChooserPromise = page.waitForEvent("filechooser");
await page.locator('[data-action="upload"][data-context="evaluation-dataset"]').click();
const invalidDatasetChooser = await invalidDatasetChooserPromise;
await invalidDatasetChooser.setFiles({
  name: "wrong-schema.jsonl",
  mimeType: "application/x-ndjson",
  buffer: Buffer.from('{"foo":"bar"}\n{"foo":"baz"}'),
});
await page.waitForTimeout(30);
await page.locator('[data-action="confirmModal"]').click();
const customDatasetCountBefore = await page.locator("#customDatasetRows tr").count();
await page.locator('[data-action="saveDatasetUpload"]').click();
check((await page.locator("#customDatasetRows tr").count()) === customDatasetCountBefore, "custom dataset accepted records missing mapped Schema fields");
check(await page.locator("#datasetUploadPanel").isVisible(), "Schema failure closed the dataset editor instead of allowing correction");
const validDatasetChooserPromise = page.waitForEvent("filechooser");
await page.locator('[data-action="upload"][data-context="evaluation-dataset"]').click();
const validDatasetChooser = await validDatasetChooserPromise;
await validDatasetChooser.setFiles({
  name: "valid-schema.jsonl",
  mimeType: "application/x-ndjson",
  buffer: Buffer.from('{"prompt":"Q1","label":"A1"}\n{"prompt":"Q2","label":"A2"}'),
});
await page.waitForTimeout(30);
await page.locator('[data-action="confirmModal"]').click();
await page.locator('[data-action="saveDatasetUpload"]').click();
check((await page.locator("#customDatasetRows tr").count()) === customDatasetCountBefore + 1, "custom dataset with valid mapped fields was not saved");
check(!(await page.locator("#datasetUploadPanel").isVisible()), "valid dataset save left editor open");
await page.locator('[data-section="5"]').click();
check((await page.locator(".metric-guide").count()) >= 3, "metric library lacks explanation/formula/range rows");

// Delivery and service management are visible, editable workflows.
await page.locator('[data-page="inference"]').first().click();
await page.locator('[data-section="0"]').click();
check((await page.locator("#modelScenario").count()) === 1 && (await page.locator("#modelModality").count()) === 1 && (await page.locator("#modelCreator").count()) === 1 && (await page.locator("#modelSort").count()) === 1, "inference model filter dimensions missing");
check((await page.locator("#modelLanguage").count()) === 1 && (await page.locator("#modelCategory").count()) === 1, "inference model language/category filters missing");
await page.locator('[data-action="generateSelectedModels"]').click();
check((await page.locator("#parallelOutputs .model-card").count()) === 2, "2–4 model experience did not use selected models");
check((await page.locator("#experienceLogRows tr").first().innerText()).includes("temperature"), "experience log lacks full parameters");
await page.locator('[data-section="1"]').click();
check((await page.locator("#calibrationDataset").count()) === 1 && (await page.locator("#compressionAdvanced").count()) === 1, "delivery compression parameters missing");
check((await page.locator("#deliveryEvalDataset").count()) === 1 && (await page.locator("#containerLimits").count()) === 1, "delivery evaluation dataset or CPU/memory limits missing");
check((await page.locator("#deliveryTasks tr").first().innerText()).includes("P95 延迟") && (await page.locator("#deliveryTasks tr").first().innerText()).includes("Accuracy"), "compression result lacks before/after latency or accuracy");
check((await page.locator("#deliveryQualityEvidence").innerText()).includes("Precision–Recall") && (await page.locator("#deliveryQualityEvidence").innerText()).includes("混淆矩阵"), "delivery quality evidence lacks PR curve or confusion matrix");
await page.locator('[data-action="monitorService"]').click();
check(await page.locator("#deliveryMonitorPanel").isVisible(), "delivery monitoring panel did not open visibly");
check((await page.locator("#deliveryMonitorPanel .metric").count()) === 4, "delivery monitoring lacks resource/QPS/error metrics");
check((await page.locator('[data-action="downloadDeliveryLogs"]').count()) === 1, "delivery runtime logs lack a download action");
await page.screenshot({ path: screenshotPath("delivery-quality-desktop.png"), fullPage: true });
await page.locator('[data-section="2"]').click();
check((await page.locator("#serviceOpsCharts svg").count()) === 3, "service operations lacks pie/bar/trend visualizations");
await page.screenshot({ path: screenshotPath("service-operations-desktop.png"), fullPage: true });
await page.locator('[data-route-action="version"]').first().click();
check((await page.locator("#routeRows tr").first().innerText()).includes("v4"), "route version did not increment from v3 to v4");
check(!(await page.locator("#routeRows tr").first().innerText()).includes("vNaN"), "route version became vNaN");
await page.locator('[data-action="editCredential"]').first().click();
await page.locator("#credentialScopeEdit").fill("only.audit.read");
await page.locator('[data-action="confirmModal"]').click();
check((await page.locator("#credentialRows tr").first().innerText()).includes("only.audit.read"), "credential permission edit did not persist in the list");
const routeCountBefore = await page.locator("#routeRows tr").count();
await page.locator('[data-action="newRoute"]').click();
await page.locator("#apiRoute").fill("/v1/qa-route");
await page.locator('[data-action="saveRoute"]').click();
check((await page.locator("#routeRows tr").count()) === routeCountBefore + 1, "route create did not update table");
await page.locator('[data-route-action="edit"]').last().click();
await page.locator("#routeScope").fill("qa.invoke");
await page.locator('[data-action="saveRoute"]').click();
check((await page.locator("#routeRows").innerText()).includes("qa.invoke"), "route edit did not persist permissions");
const stepCountBefore = await page.locator("[data-orchestration-index]").count();
await page.locator('[data-action="addOrchestrationStep"]').click();
check((await page.locator("[data-orchestration-index]").count()) === stepCountBefore + 1, "orchestration step create failed");
await page.locator('[data-action="runOrchestration"]').first().click();
check((await page.locator("#serviceDetailRows").innerText()).includes("ORCH-"), "orchestration run did not create/show instances");
await page.locator('[data-action="auditLogs"]').first().click();
check((await page.locator("#serviceDetailHead").innerText()).includes("请求参数"), "service audit log lacks request parameters");

// Documentation has a real reader and broad API contract.
await page.locator('[data-page="docs"]').first().click();
await page.locator('[data-doc-section="configuration"]').first().click();
check((await page.locator("#docReader").innerText()).includes("检查点"), "document reader did not navigate to real section content");
check((await page.locator(".api-list .endpoint").count()) >= 8, "API reference lacks endpoint coverage");
check((await page.locator("#apiParams tr").count()) >= 4, "API reference lacks parameter details");
await page.locator('[data-code-lang="python"]').click();
const pythonApiCode = await page.locator("#apiCode").innerText();
check(pythonApiCode.includes("\n") && !pythonApiCode.includes("\\n"), "Python API example contains literal escaped newlines");
check(pythonApiCode.split("\n").length >= 3, "Python API example is not formatted as runnable multiline code");

// Superscale model board: isolated business IA, dependency-aware creation and 36-item acceptance traceability.
await page.locator('[data-page="superscale-models"]').first().click();
check((await page.locator('[data-module="superscale-models"]').count()) === 9, "superscale board does not expose 9 business sections");
check((await page.locator(".super-family").count()) === 5, "superscale overview does not expose five model families");
for (let index = 0; index < 9; index += 1) {
  await page.locator(`[data-module="superscale-models"][data-section="${index}"]`).click();
  check((await page.locator(".work-card h2").count()) > 0, `superscale section ${index}: no business title`);
}
await page.locator('[data-module="superscale-models"][data-section="5"]').click();
const evalRowsBefore = await page.locator(".super-table tbody tr").count();
await page.locator('[data-action="superOpenDrawer"][data-super-kind="evaluation"]').click();
check(await page.locator("#drawer").evaluate((el) => el.classList.contains("product-drawer")), "superscale evaluation did not open the product drawer");
check((await page.locator("#superEvalTask").count()) === 1 && (await page.locator("#superEvalModel").count()) === 1 && (await page.locator("#superEvalDataset").count()) === 1, "superscale evaluation lacks dependency controls");
await page.locator('[data-action="superEvalType"][data-super-family="多模态"]').click();
check((await page.locator("#superEvalTask").innerText()).includes("视觉问答"), "superscale model type did not link evaluation tasks");
check((await page.locator("#superEvalDataset").innerText()).includes("MMBench"), "superscale model type did not link datasets");
await page.locator("#superEvalName").fill("");
await page.locator('[data-action="superSubmitDrawer"]').click();
check((await page.locator("#drawerBackdrop.open").count()) === 1, "superscale required validation allowed an empty task name");
await page.locator("#superEvalName").fill("多模态综合能力验收评测");
await page.locator('[data-action="superSubmitDrawer"]').click();
check((await page.locator("#modalBackdrop.open").count()) === 1, "superscale valid evaluation did not create a task");
await page.locator('[data-action="confirmModal"]').click();
check((await page.locator(".super-table tbody tr").count()) === evalRowsBefore + 1, "superscale evaluation task was not added to the list");
await page.locator('[data-module="superscale-models"][data-section="8"]').click();
check((await page.locator("[data-super-map-row]").count()) === 36, "superscale acceptance map does not contain 36 leaf requirements");
await page.locator("#superMappingSearch").fill("1.3.2.5.2.4.3.1.5");
await page.locator('[data-action="superFilterMapping"]').click();
check((await page.locator("[data-super-map-row]").count()) === 1, "superscale clause search did not isolate one requirement");
check((await page.locator("[data-super-map-row]").innerText()).includes("多模态推理能力"), "superscale clause search returned the wrong requirement");
await page.locator('[data-action="superBusinessEntry"]').click();
check((await page.locator(".super-section-head h2").textContent()) === "能力体验", "superscale acceptance business entry did not return to the real page");
await page.locator('[data-action="superAcceptanceInspector"]').click();
check((await page.locator("#drawerBackdrop.open").count()) === 1 && (await page.locator("#drawer").innerText()).includes("多模态推理能力"), "superscale page inspector lacks linked clauses");
await page.keyboard.press("Escape");
await page.screenshot({ path: screenshotPath("superscale-desktop.png"), fullPage: true });

// Mobile navigation: scrim, focus trap, aria-expanded reset and no overflow.
await page.setViewportSize({ width: 390, height: 844 });
await page.goto(freshUrl("overview"), { waitUntil: "networkidle" });
await page.locator("#toasts").evaluate((el) => el.replaceChildren());
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
await page.waitForTimeout(300);
await page.screenshot({ path: screenshotPath("overview-mobile.png"), fullPage: true });

await page.setViewportSize({ width: 320, height: 700 });
await page.goto(freshUrl("evaluation"), { waitUntil: "networkidle" });
const widths = await page.evaluate(() => ({ scroll: document.documentElement.scrollWidth, inner: innerWidth }));
check(widths.scroll === widths.inner, `320px viewport overflow: ${JSON.stringify(widths)}`);

await page.goto(freshUrl("inference"), { waitUntil: "networkidle" });
await page.locator('[data-module="inference"][data-section="1"]').click();
const deliveryMobileWidths = await page.evaluate(() => ({ scroll: document.documentElement.scrollWidth, inner: innerWidth }));
check(deliveryMobileWidths.scroll === deliveryMobileWidths.inner, `delivery 320px viewport overflow: ${JSON.stringify(deliveryMobileWidths)}`);
await page.screenshot({ path: screenshotPath("delivery-quality-mobile.png"), fullPage: true });
await page.locator('[data-module="inference"][data-section="2"]').click();
const serviceMobileWidths = await page.evaluate(() => ({ scroll: document.documentElement.scrollWidth, inner: innerWidth }));
check(serviceMobileWidths.scroll === serviceMobileWidths.inner, `service operations 320px viewport overflow: ${JSON.stringify(serviceMobileWidths)}`);
check((await page.locator("#serviceOpsCharts .chart").count()) === 3, "service operations charts missing on mobile");
await page.screenshot({ path: screenshotPath("service-operations-mobile.png"), fullPage: true });

await page.goto(freshUrl("superscale-models"), { waitUntil: "networkidle" });
await page.locator('[data-module="superscale-models"][data-section="5"]').click();
await page.locator('[data-action="superOpenDrawer"][data-super-kind="evaluation"]').click();
await page.waitForTimeout(300);
check(await page.locator("#drawer").evaluate((el) => el.getBoundingClientRect().width === innerWidth), "superscale mobile product drawer is not full width");
const superMobileWidths = await page.evaluate(() => ({ scroll: document.documentElement.scrollWidth, inner: innerWidth }));
check(superMobileWidths.scroll === superMobileWidths.inner, `superscale 320px viewport overflow: ${JSON.stringify(superMobileWidths)}`);
await page.screenshot({ path: screenshotPath("superscale-mobile-drawer.png"), fullPage: true });
await page.keyboard.press("Escape");

// Standalone file:// execution.
const filePage = await browser.newPage({ viewport: { width: 1200, height: 800 } });
const fileErrors = [];
filePage.on("pageerror", (error) => fileErrors.push(error.message));
await filePage.goto(new URL("../index.html", import.meta.url).href, { waitUntil: "load" });
check((await filePage.locator("#nav button").count()) >= 11, "file:// standalone navigation failed");
check(fileErrors.length === 0, `file:// errors: ${fileErrors.join("; ")}`);
await filePage.close();

await browser.close();
const report = {
  moduleCount: moduleIds.length,
  sectionCount,
  productionScreensVisited,
  consoleErrors,
  failures,
  passed: failures.length === 0 && consoleErrors.length === 0,
};
console.log(JSON.stringify(report, null, 2));
if (!report.passed) process.exitCode = 1;

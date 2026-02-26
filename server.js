const express = require('express');
const PptxGenJS = require("pptxgenjs");
const OpenAI = require('openai'); // 👈 确保引入了这个库
require('dotenv').config();
const app = express();

app.use(express.json());
app.use(express.static('public'));

// 1. 配置 AI 客户端 (请确保填入你从 ChatAnywhere 拿到的 Key)
const openai = new OpenAI({
    apiKey: process.env.OPENAI_API_KEY,
    baseURL: 'https://api.chatanywhere.tech/v1'
});

function getModelCandidates() {
    const raw = String(process.env.OPENAI_MODELS || "").trim();
    if (raw) {
        return raw.split(",").map((s) => s.trim()).filter(Boolean);
    }
    return ["deepseek-v3-2-exp", "deepseek-r1", "gpt-4o-mini"];
}

async function createCompletionWithFallback(payload, modelCandidates = []) {
    const models = modelCandidates.length ? modelCandidates : getModelCandidates();
    const timeoutMs = Number(process.env.OPENAI_MODEL_TIMEOUT_MS || 45000);
    const errors = [];
    for (const model of models) {
        try {
            const completion = await Promise.race([
                openai.chat.completions.create({
                    ...payload,
                    model
                }),
                new Promise((_, reject) =>
                    setTimeout(() => reject(new Error(`timeout after ${timeoutMs}ms`)), timeoutMs)
                )
            ]);
            const content = completion?.choices?.[0]?.message?.content || "";
            if (String(content).trim()) {
                return { model, content };
            }
            errors.push(`${model}: empty content`);
        } catch (err) {
            const msg = err?.error?.message || err?.message || "unknown error";
            errors.push(`${model}: ${msg}`);
        }
    }
    const error = new Error(`all models failed: ${errors.join(" | ")}`);
    error.details = errors;
    throw error;
}

function buildLocalExpertDeck(topic = "AI 主题演示", pageCount = 12) {
    const pages = Math.max(8, Math.min(16, Number(pageCount) || 12));
    const titles = [
        "项目背景与问题定义",
        "目标设定与评估指标",
        "方案设计与技术路线",
        "核心流程与关键模块",
        "实施计划与时间里程碑",
        "资源投入与风险控制",
        "阶段结果与数据表现",
        "对比分析与优化方向",
        "落地路径与协同机制",
        "总结与下一步计划",
        "Q&A"
    ];
    const layouts = ["封面", "章节过渡", "双栏要点", "图文右", "时间线", "对比结论", "数据重点", "图文左", "双栏要点", "总结收束", "总结收束"];
    const makeBullets = (i) => [
        `围绕${topic}明确第${i + 1}阶段目标，负责人在两周内完成方案落地，达成率目标提升15%`,
        `建立周报机制跟踪关键指标与风险项，按里程碑推进执行，确保交付质量稳定可控`,
        `结合案例复盘当前瓶颈与改进空间，提出可执行动作清单并同步资源投入边界`,
        `面向评审老师与同学输出结论先行表达，量化展示投入产出比并形成闭环复盘机制`
    ];

    const deck = {
        pages: Array.from({ length: pages }).map((_, idx) => {
            if (idx === 0) {
                return normalizeVisualFields({
                    title: topic,
                    page_type: "封面",
                    layout: "封面",
                    bullets: [
                        "专家模式自动生成",
                        "结构化叙事与数据化表达",
                        "可直接下载并用于答辩展示",
                        "支持后续逐页微调与优化"
                    ],
                    visual_suggestion: "使用高对比主标题与简洁副标题，突出主题辨识度",
                    note: "开场先给结论，再说明本次汇报结构、评估维度与预期成果。",
                    should_use_icon: true,
                    should_use_chart: false,
                    should_use_big_number: false,
                    visual_priority: "high",
                    layout_density: "light"
                });
            }
            const t = titles[(idx - 1) % titles.length];
            const layout = layouts[(idx - 1) % layouts.length];
            return normalizeVisualFields({
                title: t,
                page_type: t,
                layout,
                bullets: makeBullets(idx),
                visual_suggestion: "建议配合场景图、流程图或指标卡增强说服力",
                note: "讲解时按“结论-证据-动作”顺序展开，强调时间节点与责任分工。",
                should_use_icon: true,
                should_use_chart: /数据|指标|结果|对比/.test(t),
                should_use_big_number: /结果|指标/.test(t),
                visual_priority: "high",
                layout_density: idx % 4 === 0 ? "normal" : "dense"
            });
        })
    };
    return strengthenDeckJson(deck, { topic, audience: "评审老师 + 同学", tone: "专家评审、结论先行" });
}

function normalizeImageQuery(input = "") {
    const text = String(input).toLowerCase();
    const map = [
        { keys: ["ai", "人工智能", "大模型", "算法", "机器学习"], tag: "artificial-intelligence" },
        { keys: ["教育", "学校", "课程", "答辩", "课堂", "大学"], tag: "education" },
        { keys: ["商业", "市场", "运营", "管理", "策略", "增长"], tag: "business" },
        { keys: ["数据", "图表", "指标", "分析"], tag: "data" },
        { keys: ["团队", "协作", "社团"], tag: "teamwork" },
        { keys: ["金融", "投资", "预算"], tag: "finance" },
        { keys: ["产品", "发布", "用户"], tag: "product" },
        { keys: ["科技", "数字化", "系统"], tag: "technology" }
    ];
    const hit = map.find(({ keys }) => keys.some((k) => text.includes(String(k).toLowerCase())));
    return hit ? hit.tag : "presentation";
}

function extractKeywords(input = "", limit = 3) {
    const raw = String(input || "")
        .replace(/[^\u4e00-\u9fa5A-Za-z0-9\s]/g, " ")
        .split(/\s+/)
        .map((s) => s.trim())
        .filter(Boolean);
    const stop = new Set(["the", "and", "for", "with", "from", "this", "that", "ppt", "slide", "内容", "页面"]);
    const freq = new Map();
    raw.forEach((w) => {
        if (w.length < 2 || stop.has(w.toLowerCase())) return;
        freq.set(w, (freq.get(w) || 0) + 1);
    });
    return [...freq.entries()]
        .sort((a, b) => b[1] - a[1])
        .slice(0, limit)
        .map(([k]) => k);
}

function normalizeVisualFields(page = {}) {
    const p = { ...page };
    const toBool = (v) => v === true;
    const clean = (v, fallback, allow) => (allow.includes(v) ? v : fallback);

    p.visual_priority = clean(p.visual_priority, "high", ["low", "medium", "high"]);
    p.layout_density = clean(p.layout_density, "dense", ["light", "normal", "dense"]);
    p.should_use_icon = toBool(p.should_use_icon);
    p.should_use_chart = toBool(p.should_use_chart);
    p.should_use_big_number = toBool(p.should_use_big_number);

    const typeText = String(p.page_type || p.layout || p.title || "");
    if (/实验结果/.test(typeText)) {
        p.should_use_big_number = true;
        p.visual_priority = "high";
        p.layout_style = "big_number";
    }
    if (/创新/.test(typeText)) {
        p.layout_style = "visual_focus";
        p.visual_priority = "high";
    }
    if (/总结/.test(typeText)) {
        p.layout_style = "minimal";
        p.layout_density = "light";
    }

    if (!p.should_use_icon && !p.should_use_chart && !p.should_use_big_number) {
        p.should_use_icon = true;
    }
    return p;
}

function inferExpertLayout(page = {}, idx = 0, total = 10) {
    const text = `${page.title || ""} ${page.page_type || ""} ${page.layout || ""}`.toLowerCase();
    if (idx === 0) return "封面";
    if (idx === total - 1 || /总结|致谢|q&a/.test(text)) return "总结收束";
    if (/目录|章节|议程/.test(text)) return "章节过渡";
    if (/时间|阶段|里程碑|路线|计划/.test(text)) return "时间线";
    if (/数据|指标|增长|转化|实验|结果|统计|成本|收益/.test(text)) return "数据重点";
    if (/对比|差异|方案|现状|优劣/.test(text)) return "对比结论";
    if (/案例|场景|用户|产品|demo|原型|设计/.test(text)) return idx % 2 === 0 ? "图文左" : "图文右";
    return idx % 3 === 0 ? "双栏要点" : "图文右";
}

function parseDeckSections(input = "") {
    const raw = unwrapJsonText(input);
    if (raw.startsWith("{")) {
        try {
            const obj = JSON.parse(raw);
            const pages = Array.isArray(obj.pages) ? obj.pages : [];
            return pages.map((p) => {
                const page = normalizeVisualFields(p || {});
                return {
                    title: String(page.title || "未命名页面"),
                    layout: String(page.layout || page.layout_style || "双栏要点"),
                    visual: String(page.visual_suggestion || page.visual || ""),
                    note: String(page.note || ""),
                    bullets: Array.isArray(page.bullets) ? page.bullets.map((b) => String(b)).filter(Boolean) : [],
                    page_type: String(page.page_type || ""),
                    visual_priority: page.visual_priority,
                    should_use_icon: page.should_use_icon,
                    should_use_chart: page.should_use_chart,
                    should_use_big_number: page.should_use_big_number,
                    layout_density: page.layout_density,
                    layout_style: String(page.layout_style || "")
                };
            });
        } catch (_) {
            // ignore and fallback to markdown parsing
        }
    }

    return String(input)
        .split("##")
        .map((s) => s.trim())
        .filter(Boolean)
        .map((section) => {
            const lines = section.split('\n').map((l) => l.trim()).filter(Boolean);
            const title = (lines[0] || "未命名页面").replace(/^#+\s*/, '').trim();
            const layoutLine = lines.find((l) => l.startsWith('[版式]'));
            const visualLine = lines.find((l) => l.startsWith('[视觉建议]'));
            const noteLine = lines.find((l) => l.startsWith('[备注]'));
            const bullets = lines
                .filter((l) => l.startsWith('- '))
                .map((l) => l.replace(/^- /, '').trim())
                .filter(Boolean);

            return {
                title,
                layout: layoutLine ? layoutLine.replace('[版式]：', '').trim() : "双栏要点",
                visual: visualLine ? visualLine.replace('[视觉建议]：', '').trim() : "",
                note: noteLine ? noteLine.replace('[备注]：', '').trim() : "",
                bullets
            };
        });
}

function unwrapJsonText(raw = "") {
    const text = String(raw || "").trim();
    if (!text) return "";
    const noFence = text
        .replace(/^```json\s*/i, "")
        .replace(/^```\s*/i, "")
        .replace(/\s*```$/i, "")
        .trim();
    const start = noFence.indexOf("{");
    const end = noFence.lastIndexOf("}");
    if (start >= 0 && end > start) return noFence.slice(start, end + 1);
    return noFence;
}

function normalizeDeckJson(raw = "") {
    try {
        const obj = JSON.parse(unwrapJsonText(raw));
        const pages = Array.isArray(obj.pages) ? obj.pages : [];
        return {
            pages: pages.map((p) => normalizeVisualFields({
                title: p?.title || "未命名页面",
                page_type: p?.page_type || p?.layout || "",
                layout: p?.layout || "双栏要点",
                layout_style: p?.layout_style || "",
                bullets: Array.isArray(p?.bullets) ? p.bullets : [],
                visual_suggestion: p?.visual_suggestion || p?.visual || "",
                note: p?.note || "",
                visual_priority: p?.visual_priority,
                should_use_icon: p?.should_use_icon,
                should_use_chart: p?.should_use_chart,
                should_use_big_number: p?.should_use_big_number,
                layout_density: p?.layout_density
            }))
        };
    } catch (_) {
        return null;
    }
}

function sectionsToDeckJson(sections = []) {
    return {
        pages: sections.map((s) => {
            const page = normalizeVisualFields({
                title: s?.title || "未命名页面",
                page_type: s?.page_type || s?.layout || "",
                layout: s?.layout || "双栏要点",
                layout_style: s?.layout_style || "",
                bullets: Array.isArray(s?.bullets) ? s.bullets : [],
                visual_suggestion: s?.visual || "",
                note: s?.note || "",
                visual_priority: s?.visual_priority,
                should_use_icon: s?.should_use_icon,
                should_use_chart: s?.should_use_chart,
                should_use_big_number: s?.should_use_big_number,
                layout_density: s?.layout_density
            });
            return {
                title: page.title,
                page_type: page.page_type,
                layout: page.layout,
                layout_style: page.layout_style || "",
                bullets: page.bullets,
                visual_suggestion: page.visual_suggestion,
                note: page.note,
                visual_priority: page.visual_priority,
                should_use_icon: page.should_use_icon,
                should_use_chart: page.should_use_chart,
                should_use_big_number: page.should_use_big_number,
                layout_density: page.layout_density
            };
        })
    };
}

function enrichBulletText(text = "", idx = 0) {
    let t = String(text || "").replace(/\s+/g, " ").trim();
    if (!t) return "";

    if (t.length < 16) {
        const tails = [
            "，结合现状给出可执行方案与负责人",
            "，明确时间节点并设置阶段验收指标",
            "，补充量化目标与资源投入边界",
            "，对应关键风险并给出应对策略"
        ];
        t += tails[idx % tails.length];
    }
    if (!/(%|倍|人|项|万元|小时|天|周|月|学期|季度|\d)/.test(t)) {
        t += "，目标指标提升15%";
    }
    if (!/(负责|执行|落地|推进|优化|建立|复盘|跟踪)/.test(t)) {
        t += "，并安排执行与复盘机制";
    }

    if (t.length > 42) t = t.slice(0, 42);
    return t;
}

function strengthenDeckJson(deck, context = {}) {
    const topic = String(context.topic || "主题").trim();
    const audience = String(context.audience || "听众").trim();
    const tone = String(context.tone || "专家评审语气").trim();
    const pages = Array.isArray(deck?.pages) ? deck.pages : [];

    return {
        pages: pages.map((p, pIdx) => {
            const normalized = normalizeVisualFields(p || {});
            let bullets = Array.isArray(normalized.bullets)
                ? normalized.bullets.map((b, idx) => enrichBulletText(b, idx)).filter(Boolean)
                : [];

            if (bullets.length < 4) {
                const fillers = [
                    `围绕${topic}拆解当前问题、目标与优先级，形成执行清单`,
                    `针对${audience}优化表达方式，确保结论可理解可落地`,
                    `按${tone}语气输出关键结论，并标注阶段里程碑`,
                    `建立数据看板，按周追踪指标变化并持续复盘优化`
                ].map((b, idx) => enrichBulletText(b, idx));
                bullets = [...bullets, ...fillers].slice(0, 6);
            } else if (bullets.length > 6) {
                bullets = bullets.slice(0, 6);
            }

            let note = String(normalized.note || "").trim();
            if (note.length < 24) {
                note = `讲解建议：先用一句话交代本页结论，再说明关键数据来源、执行路径与风险对策，最后强调下一阶段里程碑与责任分工。`;
            }
            if (note.length > 90) note = note.slice(0, 90);

            const autoLayout = inferExpertLayout(normalized, pIdx, pages.length);
            const hasDataSignal = bullets.some((b) => /%|倍|人|项|万元|小时|天|周|月|学期|季度|\d/.test(b));
            const shouldBigNumber = normalized.should_use_big_number || /实验|结果|增长|转化|指标/.test(`${normalized.title}${normalized.page_type}`);
            const shouldChart = normalized.should_use_chart || hasDataSignal || autoLayout === "数据重点";
            const shouldIcon = normalized.should_use_icon || (!shouldChart && !shouldBigNumber);

            return {
                ...normalized,
                layout: autoLayout,
                bullets,
                note,
                visual_priority: "high",
                layout_density: normalized.layout_style === "minimal" ? "light" : "dense",
                should_use_icon: shouldIcon,
                should_use_chart: shouldChart,
                should_use_big_number: shouldBigNumber
            };
        })
    };
}

function densifySections(sections = []) {
    const out = [];
    for (const section of sections) {
        const textLen = `${section.title}${section.visual}${(section.bullets || []).join('')}`.replace(/\s/g, '').length;
        const forceImage = textLen < 50;
        const n = (section.bullets || []).length;
        if (n <= 5) {
            out.push({ ...section, _forceImage: forceImage });
            continue;
        }
        if (n === 6) {
            out.push(
                { ...section, bullets: section.bullets.slice(0, 3), layout: "紧凑要点", _forceImage: false },
                { ...section, bullets: section.bullets.slice(3), layout: "紧凑要点", title: `${section.title}（续1）`, _forceImage: false }
            );
            continue;
        }
        out.push({ ...section, bullets: section.bullets.slice(0, 4), layout: "2x2宫格", _forceImage: false });
        const rest = section.bullets.slice(4);
        for (let i = 0; i < rest.length; i += 4) {
            const group = rest.slice(i, i + 4);
            out.push({
                ...section,
                bullets: group,
                layout: group.length <= 3 ? "紧凑要点" : "2x2宫格",
                title: `${section.title}（续${Math.floor(i / 4) + 1}）`,
                _forceImage: false
            });
        }
    }
    return out;
}

function getPptTheme(style) {
    const packs = {
        "科技感": { bg: "0A1224", primary: "F4F8FF", secondary: "BFD1ED", accent: "38BDF8" },
        "商务简约": { bg: "FAF8F4", primary: "1D2430", secondary: "526071", accent: "0F766E" },
        "课程汇报": { bg: "F4F8FF", primary: "1A315E", secondary: "4A6BA6", accent: "2A8CFF" },
        "答辩展示": { bg: "F8F6FF", primary: "2D2A59", secondary: "62609A", accent: "6B7CFF" },
        "社团活动": { bg: "FFF7F2", primary: "5B2D1D", secondary: "94614D", accent: "FF8A5B" },
        "竞赛路演": { bg: "F2FFFA", primary: "134236", secondary: "3D7669", accent: "1EBE9D" },
        "专家模式": { bg: "0D1020", primary: "F5F8FF", secondary: "BAC7E8", accent: "56A8FF" }
    };
    return packs[style] || packs["专家模式"];
}

async function fetchImageDataUri(query = "") {
    const keyword = normalizeImageQuery(query);
    const seed = encodeURIComponent((query || keyword).slice(0, 60));
    const candidates = [
        `https://loremflickr.com/1600/900/${encodeURIComponent(keyword)}`,
        `https://source.unsplash.com/1600x900/?${encodeURIComponent(keyword)}`,
        `https://picsum.photos/seed/${seed}/1600/900`
    ];

    for (const url of candidates) {
        try {
            const controller = new AbortController();
            const timer = setTimeout(() => controller.abort(), 2500);
            const resp = await fetch(url, { signal: controller.signal, redirect: 'follow' });
            clearTimeout(timer);
            const type = resp.headers.get('content-type') || '';
            if (!resp.ok || !type.startsWith('image/')) continue;
            const arrBuf = await resp.arrayBuffer();
            const b64 = Buffer.from(arrBuf).toString('base64');
            return `data:${type};base64,${b64}`;
        } catch (_) {
            continue;
        }
    }
    return null;
}

function fallbackSvgDataUri(section, theme, query = "") {
    const esc = (s = "") => String(s).replace(/[<>&"]/g, '').slice(0, 40);
    const title = esc(section.title || "AI Presentation");
    const sub = esc(section.visual || section.bullets?.[0] || "Insight");
    const category = normalizeImageQuery(query || `${section.title} ${section.visual}`);
    const iconMap = {
        "artificial-intelligence": "AI",
        "education": "EDU",
        "business": "BIZ",
        "data": "DATA",
        "teamwork": "TEAM",
        "finance": "FIN",
        "product": "PROD",
        "technology": "TECH",
        "presentation": "IDEA"
    };
    const badge = iconMap[category] || "IDEA";
    const kws = extractKeywords(query || `${section.title} ${section.visual} ${(section.bullets || []).join(" ")}`, 3);
    const [k1 = "主题", k2 = "分析", k3 = "方案"] = kws.map(esc);
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="1600" height="900" viewBox="0 0 1600 900"><rect width="1600" height="900" fill="#${theme.bg}"/><circle cx="1280" cy="120" r="340" fill="#${theme.accent}" fill-opacity="0.28"/><circle cx="200" cy="760" r="280" fill="#${theme.secondary}" fill-opacity="0.22"/><rect x="120" y="120" width="1360" height="660" rx="34" fill="#ffffff" fill-opacity="0.08" stroke="#${theme.accent}" stroke-opacity="0.55"/><rect x="180" y="190" width="210" height="84" rx="18" fill="#${theme.accent}" fill-opacity="0.88"/><text x="285" y="245" text-anchor="middle" fill="#ffffff" font-size="38" font-family="Segoe UI, Arial" font-weight="700">${badge}</text><text x="190" y="380" fill="#${theme.accent}" font-size="84" font-family="Segoe UI, Arial" font-weight="700">${title}</text><text x="196" y="456" fill="#${theme.secondary}" font-size="42" font-family="Segoe UI, Arial">${sub}</text><rect x="190" y="520" width="220" height="56" rx="14" fill="#ffffff" fill-opacity="0.16" stroke="#${theme.accent}" stroke-opacity="0.55"/><text x="300" y="557" text-anchor="middle" fill="#${theme.accent}" font-size="28" font-family="Segoe UI, Arial">${k1}</text><rect x="430" y="520" width="220" height="56" rx="14" fill="#ffffff" fill-opacity="0.16" stroke="#${theme.accent}" stroke-opacity="0.55"/><text x="540" y="557" text-anchor="middle" fill="#${theme.accent}" font-size="28" font-family="Segoe UI, Arial">${k2}</text><rect x="670" y="520" width="220" height="56" rx="14" fill="#ffffff" fill-opacity="0.16" stroke="#${theme.accent}" stroke-opacity="0.55"/><text x="780" y="557" text-anchor="middle" fill="#${theme.accent}" font-size="28" font-family="Segoe UI, Arial">${k3}</text></svg>`;
    return `data:image/svg+xml;base64,${Buffer.from(svg).toString('base64')}`;
}

app.post('/generate', async (req, res) => {
    const { topic, type: inputType, audience: inputAudience, tone: inputTone, pageCount, skills } = req.body;
    const type = String(inputType || "答辩展示").trim();
    const audience = String(inputAudience || "评审老师 + 同学").trim();
    const tone = String(inputTone || "专家评审、结论先行").trim();
    console.log(`收到请求: 主题[${topic}], 风格[${type}], 听众[${audience}], 语气[${tone}], 页数[${pageCount}], 技能[${Array.isArray(skills) ? skills.join(',') : ''}]`);

    if (!process.env.OPENAI_API_KEY) {
        return res.status(500).json({ error: "缺少 OPENAI_API_KEY，请检查 .env 配置" });
    }

    try {
        const targetPages = Number(pageCount);
        const pages = Number.isFinite(targetPages) ? Math.min(16, Math.max(8, targetPages)) : 12;
        const tokenBudget = Math.min(3600, Math.max(1800, pages * 220));
        const skillList = Array.isArray(skills) ? skills.filter(Boolean).slice(0, 6) : [];
        const skillMap = {
            "结构化叙事": "章节之间保持“问题-分析-方案-落地”递进，每章开头有一句过渡语。",
            "数据指标强化": "每页至少出现一个关键数字、百分比或可量化指标。",
            "案例驱动": "至少 3 页要点引用具体案例或场景，避免泛泛描述。",
            "金句总结": "每页末尾增加一句短金句，用于口播收束。",
            "行动清单": "至少 2 页提供可执行动作清单（谁、何时、做什么）。",
            "风险与对策": "至少 2 页增加风险提示及对应应对策略。"
        };
        const skillRules = skillList.map((s) => `- ${skillMap[s] || `${s}：请在内容中体现该能力`}`).join('\n');

        const pageRule = Number.isFinite(targetPages) && targetPages >= 10 && targetPages <= 16
            ? `篇幅：${targetPages}-${Math.min(targetPages + 1, 16)} 页（必须）。`
            : `篇幅：12-14 页（必须）。`;
        const defenseStructureHint = `\n- 请采用论文/项目答辩常见结构组织整套内容，建议包含但不限于：封面、目录/议程、研究背景与意义、研究/方案设计、方法与技术路线、实验设计与数据结果、结论与创新点、存在问题与不足、改进方向与后续计划、总结与致谢、Q&A。`;

        const baseMessages = [
            {
                role: "system",
                content: `你是资深的演示设计总监，请生成“适合直接排版为高质量 PPT”的完整 JSON。

硬性要求：
1. ${pageRule}
2. 输出格式必须严格为：
{
  "pages": [
    {
      "title": "页面标题",
      "page_type": "页面类型",
      "layout": "版式名",
      "layout_style": "可选：visual_focus|minimal|big_number|standard",
      "bullets": ["要点1","要点2","要点3","要点4"],
      "visual_suggestion": "一句话视觉建议",
      "note": "70-110字备注",
      "visual_priority": "low|medium|high",
      "should_use_icon": true,
      "should_use_chart": false,
      "should_use_big_number": false,
      "layout_density": "light|normal|dense"
    }
  ]
}
3. 版式名(layout)只能从以下中选择：
   - 封面
   - 章节过渡
   - 双栏要点
   - 图文左
   - 图文右
   - 时间线
   - 数据重点
   - 对比结论
   - 总结收束
4. 全文至少覆盖 7 种不同版式，且必须包含“时间线、数据重点、对比结论、图文左/图文右”。
5. 要点要求：
   - 每页 4-6 条
   - 每条 22-40 字
   - 必须包含可执行动作、数字或案例，不要空话套话。
6. 视觉规则（必须满足）：
   - 实验结果页必须 should_use_big_number=true 且 layout_style=big_number
   - 创新页必须 layout_style=visual_focus
   - 总结页必须 layout_style=minimal 且 layout_density=light
   - 每页至少一个视觉元素（icon/chart/big_number 至少一个为 true）
7. 语言必须体现专家评审水平：结论先行、证据支撑、动作闭环。
8. 只输出合法 JSON，不要解释、不要 markdown 代码块。`
            },
            {
                role: "user",
                content: `请为主题《${topic}》生成一份风格为“${type}”的高质量完整 PPT 稿件。
补充要求：
- 听众对象：${audience}
- 表达语气：${tone}
- 尽量给出可执行动作、时间节点、关键指标。
- 默认使用“背景-问题-方法-结果-行动”叙事骨架组织每页要点。
- 关键结论必须数字化表达，优先使用同比、环比、达成率、投入产出比等。
- 启用技能：${skillList.length ? skillList.join('、') : "结构化叙事、数据指标强化"}
${skillRules ? `- 额外技能规则：\n${skillRules}` : ''}${defenseStructureHint}`
            }
        ];

        // 主请求：多模型回退，减少“单模型不可用”导致的全量失败
        const primary = await createCompletionWithFallback({
            messages: baseMessages,
            max_tokens: tokenBudget,
            temperature: 0.65
        });

        let aiResult = String(primary.content || "").trim();
        let normalized = normalizeDeckJson(aiResult);

        // 若模型返回了接近 JSON 但不完整，尝试一次“修复 JSON”而非直接报错
        if (!normalized && (aiResult.startsWith("{") || aiResult.startsWith("```json"))) {
            const fixPrompt = [
                {
                    role: "system",
                    content: "你是 JSON 修复器。请把输入修复为合法 JSON，仅输出 JSON。禁止解释。"
                },
                {
                    role: "user",
                    content: `请修复这段 PPT JSON，必须保留 pages 数组结构：\n${aiResult}`
                }
            ];
            try {
                const fixed = await createCompletionWithFallback({
                    messages: fixPrompt,
                    max_tokens: Math.max(1200, Math.floor(tokenBudget * 0.7)),
                    temperature: 0.2
                }, ["gpt-5.2"]);
                aiResult = String(fixed.content || "").trim();
                normalized = normalizeDeckJson(aiResult);
            } catch (_) {
                // ignore and fallback below
            }
        }

        if (normalized) {
            const strengthened = strengthenDeckJson(normalized, { topic, audience, tone });
            return res.json({
                result: JSON.stringify(strengthened, null, 2),
                format: "json"
            });
        }

        // 最后兜底：尝试按 markdown/半结构内容解析并转成标准 JSON，避免前端空白
        const parsedSections = parseDeckSections(aiResult);
        if (parsedSections.length) {
            const fallbackDeck = strengthenDeckJson(
                sectionsToDeckJson(parsedSections),
                { topic, audience, tone }
            );
            return res.json({
                result: JSON.stringify(fallbackDeck, null, 2),
                format: "json",
                warning: "AI 返回格式异常，已自动修复为可用结构。"
            });
        }

        return res.status(502).json({
            error: "AI 返回内容无法解析为 PPT 结构，请重试（建议减少页数或切换风格）"
        });

    } catch (error) {
        console.error("AI 接口报错:", error);
        // 超时或网络抖动时：直接返回本地专家版，保证“始终可生成”
        const localDeck = buildLocalExpertDeck(topic, pageCount);
        return res.json({
            result: JSON.stringify(localDeck, null, 2),
            format: "json",
            warning: "云端模型超时，已切换为本地专家模板生成。"
        });
    }
});

app.post('/refine-item', async (req, res) => {
    const { text, mode, topic, audience, tone } = req.body;

    if (!process.env.OPENAI_API_KEY) {
        return res.status(500).json({ error: "缺少 OPENAI_API_KEY，请检查 .env 配置" });
    }
    if (!text || typeof text !== 'string') {
        return res.status(400).json({ error: "缺少有效的 text 字段" });
    }

    const normalizedMode = mode === 'expand' ? 'expand' : 'condense';
    const modeInstruction = normalizedMode === 'expand'
        ? "请在保留原意基础上扩充为更具体、可执行、有数据感的一条要点。"
        : "请提炼为更短更有力的一条要点，保留关键结论和动作。";

    try {
        const completion = await openai.chat.completions.create({
            model: "deepseek-r1",
            messages: [
                {
                    role: "system",
                    content: `你是资深演示顾问。请改写一条 PPT 要点。
要求：
1. 只输出改写后的单条文本，不要解释，不要加序号。
2. 字数控制在 18-36 字之间。
3. 语气风格与场景一致，避免空话。`
                },
                {
                    role: "user",
                    content: `主题：${topic || "未提供"}
听众：${audience || "通用听众"}
语气：${tone || "简洁专业"}
模式：${normalizedMode}
原始要点：${text}
${modeInstruction}`
                }
            ],
            max_tokens: 120,
            temperature: normalizedMode === 'expand' ? 0.75 : 0.5
        });

        const refined = completion?.choices?.[0]?.message?.content?.trim();
        if (!refined) {
            return res.status(500).json({ error: "AI 未返回有效内容" });
        }
        res.json({ result: refined.replace(/^[-*\d.\s]+/, '').trim() });
    } catch (error) {
        console.error("refine-item 接口报错:", error);
        res.status(500).json({ error: "要点改写失败，请稍后重试" });
    }
});

app.post('/export-ppt', async (req, res) => {
    const { content, style, topic, imageMode } = req.body || {};
    if (!content || typeof content !== 'string') {
        return res.status(400).json({ error: "缺少有效 content" });
    }
    try {
        const expertDeck = strengthenDeckJson(sectionsToDeckJson(parseDeckSections(content)), {
            topic,
            audience: "评审老师 + 同学",
            tone: "专家评审、结论先行"
        });
        const sections = densifySections(parseDeckSections(JSON.stringify(expertDeck)));
        if (!sections.length) {
            return res.status(400).json({ error: "未解析到可导出的页面" });
        }
        const theme = getPptTheme(style || "专家模式");
        const FONT = "Microsoft YaHei";
        const titleTopic = String(topic || "AI 主题演示").trim();

        const pptx = new PptxGenJS();
        pptx.layout = "LAYOUT_WIDE";
        pptx.author = "AI PPT Studio";
        pptx.company = "AI PPT Studio";
        pptx.subject = titleTopic;
        pptx.title = `${titleTopic} - AI Deck`;

        const useRemoteImages = imageMode === 'quality';
        const imageTasks = sections.map(async (section, idx) => {
            const layout = String(section.layout || "").toLowerCase();
            const needsImage = section._forceImage || layout.includes('图文左') || layout.includes('图文右');
            if (!needsImage) return [idx, null];
            const query = [section.title, section.visual, (section.bullets || []).slice(0, 2).join(' ')].join(' ');
            const data = useRemoteImages
                ? (await fetchImageDataUri(query) || fallbackSvgDataUri(section, theme, query))
                : fallbackSvgDataUri(section, theme, query);
            return [idx, data];
        });
        const imageMap = new Map(await Promise.all(imageTasks));

        const addBackdrop = (slide) => {
            slide.background = { fill: theme.bg };
            slide.addShape(pptx.ShapeType.ellipse, {
                x: 9.3, y: -1.2, w: 5.8, h: 5.8,
                fill: { color: theme.accent, transparency: 86 },
                line: { color: theme.accent, transparency: 100 }
            });
            slide.addShape(pptx.ShapeType.ellipse, {
                x: -1.6, y: 5.0, w: 4.8, h: 4.8,
                fill: { color: theme.secondary, transparency: 90 },
                line: { color: theme.secondary, transparency: 100 }
            });
            slide.addShape(pptx.ShapeType.rect, {
                x: 0.16, y: 0.14, w: 13.01, h: 7.2,
                fill: { color: theme.bg, transparency: 100 },
                line: { color: theme.secondary, pt: 1 }
            });
            slide.addShape(pptx.ShapeType.rect, {
                x: 0.16, y: 0.14, w: 0.07, h: 7.2,
                fill: { color: theme.accent },
                line: { color: theme.accent, transparency: 100 }
            });
        };

        const addHeader = (slide, title, section = {}) => {
            addBackdrop(slide);
            slide.addShape(pptx.ShapeType.roundRect, {
                x: 0.64, y: 0.62, w: 11.95, h: 1.34, rectRadius: 0.08,
                fill: { color: theme.secondary, transparency: 88 },
                line: { color: theme.accent, pt: 0.8, transparency: 40 }
            });
            slide.addText(title, {
                x: 0.7, y: 0.82, w: 11.8, h: 0.7,
                fontSize: 30, bold: true, color: theme.primary, fontFace: FONT
            });
            slide.addText(titleTopic, {
                x: 0.72, y: 1.42, w: 7.2, h: 0.3,
                fontSize: 12, color: theme.secondary, fontFace: FONT
            });
            slide.addShape(pptx.ShapeType.rect, {
                x: 0.7, y: 1.72, w: 2.2, h: 0.04,
                fill: { color: theme.accent }, line: { color: theme.accent, transparency: 100 }
            });
            const tag = String(section.page_type || section.layout || "核心页面").slice(0, 14);
            slide.addShape(pptx.ShapeType.roundRect, {
                x: 10.55, y: 1.45, w: 1.95, h: 0.34, rectRadius: 0.12,
                fill: { color: theme.accent, transparency: 26 },
                line: { color: theme.accent, pt: 0.8, transparency: 30 }
            });
            slide.addText(tag, {
                x: 10.62, y: 1.51, w: 1.8, h: 0.22,
                align: 'center', fontSize: 9, bold: true, color: theme.primary, fontFace: FONT
            });
        };

        const addFooter = (slide, idx, total) => {
            slide.addShape(pptx.ShapeType.line, {
                x: 0.72, y: 6.78, w: 11.88, h: 0,
                line: { color: theme.secondary, pt: 0.8, transparency: 45 }
            });
            slide.addText(`${String(idx + 1).padStart(2, "0")} / ${String(total).padStart(2, "0")}`, {
                x: 11.35, y: 6.84, w: 1.3, h: 0.24,
                align: 'right', fontSize: 10, color: theme.accent, fontFace: FONT
            });
        };

        const addBulletList = (slide, bullets, x, y, w, h, fontSize = 15) => {
            if (!bullets?.length) return;
            const lineH = Math.min(0.8, Math.max(0.54, h / bullets.length));
            bullets.forEach((text, i) => {
                const yy = y + i * lineH;
                if (yy + lineH > y + h) return;
                slide.addShape(pptx.ShapeType.roundRect, {
                    x, y: yy + 0.08, w, h: Math.max(0.36, lineH - 0.14), rectRadius: 0.05,
                    fill: { color: theme.secondary, transparency: 90 },
                    line: { color: theme.accent, pt: 0.4, transparency: 65 }
                });
                slide.addText([
                    { text: "▸ ", options: { color: theme.accent, bold: true, fontSize: fontSize + 1, fontFace: FONT } },
                    { text, options: { color: theme.secondary, fontSize, fontFace: FONT } }
                ], { x, y: yy, w, h: lineH, valign: "mid" });
            });
        };

        const extractBigNumber = (section = {}) => {
            const text = `${section.visual || ""} ${(section.bullets || []).join(" ")}`;
            const hit = text.match(/(\d+(?:\.\d+)?\s*(?:%|倍|项|人|万元|亿|天|周|月)?)/);
            return hit ? hit[1] : "15%";
        };

        const addBigNumberKpi = (slide, section) => {
            if (!section.should_use_big_number) return;
            const kpi = extractBigNumber(section);
            slide.addShape(pptx.ShapeType.roundRect, {
                x: 10.35, y: 5.75, w: 2.25, h: 0.95, rectRadius: 0.12,
                fill: { color: theme.accent, transparency: 22 },
                line: { color: theme.accent, pt: 1.1 }
            });
            slide.addText(kpi, {
                x: 10.45, y: 5.89, w: 2.0, h: 0.38,
                align: "center", fontSize: 24, bold: true, color: theme.primary, fontFace: FONT
            });
            slide.addText("关键指标", {
                x: 10.45, y: 6.31, w: 2.0, h: 0.2,
                align: "center", fontSize: 9, color: theme.primary, fontFace: FONT
            });
        };

        const addVisualBadge = (slide, section) => {
            if (!section.should_use_icon) return;
            slide.addShape(pptx.ShapeType.roundRect, {
                x: 11.52, y: 0.72, w: 1.06, h: 0.34, rectRadius: 0.1,
                fill: { color: theme.accent, transparency: 20 },
                line: { color: theme.accent, pt: 0.8 }
            });
            slide.addText("EXPERT", {
                x: 11.58, y: 0.78, w: 0.95, h: 0.2, align: "center",
                fontSize: 9, bold: true, color: theme.primary, fontFace: FONT
            });
        };

        const addMiniChart = (slide, section) => {
            if (!section.should_use_chart) return;
            const values = (section.bullets || []).slice(0, 4).map((t, idx) => {
                const m = String(t).match(/(\d+(?:\.\d+)?)/);
                const n = m ? Number(m[1]) : 40 + idx * 12;
                return Math.max(18, Math.min(95, n));
            });
            values.forEach((v, idx) => {
                const x = 9.2 + idx * 0.78;
                const h = (v / 100) * 1.1 + 0.2;
                slide.addShape(pptx.ShapeType.roundRect, {
                    x, y: 5.9 - h, w: 0.5, h, rectRadius: 0.04,
                    fill: { color: theme.accent, transparency: 24 },
                    line: { color: theme.accent, transparency: 100 }
                });
            });
        };

        const addSlideWithImage = (slide, section, imageData, imageBox, bodyBox, bodyText, bodyAlign = "left", bodyFontSize = 14) => {
            addHeader(slide, section.title, section);
            slide.addShape(pptx.ShapeType.roundRect, {
                x: imageBox.x, y: imageBox.y, w: imageBox.w, h: imageBox.h,
                rectRadius: 0.1, fill: { color: theme.secondary, transparency: 86 }, line: { color: theme.accent, pt: 1.2 }
            });
            if (imageData) {
                slide.addImage({ data: imageData, x: imageBox.x + 0.05, y: imageBox.y + 0.05, w: imageBox.w - 0.1, h: imageBox.h - 0.1 });
            }
            slide.addText(bodyText, {
                x: bodyBox.x, y: bodyBox.y, w: bodyBox.w, h: bodyBox.h,
                align: bodyAlign, fontSize: bodyFontSize, color: theme.secondary, fontFace: FONT
            });
        };

        for (let i = 0; i < sections.length; i += 1) {
            const section = sections[i];
            const layout = String(section.layout || "").toLowerCase();
            const pageType = String(section.page_type || "").toLowerCase();
            const slide = pptx.addSlide();
            const imageData = imageMap.get(i) || null;

            if (layout.includes('封面')) {
                addBackdrop(slide);
                slide.addText(section.title, {
                    x: 0.82, y: 2.0, w: 11.7, h: 1.2, align: 'center',
                    fontSize: 44, bold: true, color: theme.primary, fontFace: FONT
                });
                slide.addText(section.visual || "AI 自动生成 · 智能排版 · 结构化表达", {
                    x: 1.6, y: 3.55, w: 10.1, h: 0.7, align: 'center',
                    fontSize: 18, color: theme.secondary, fontFace: FONT
                });
            } else if (layout.includes('章节过渡') || pageType.includes('章节') || pageType.includes('目录')) {
                // 章节封面 / 目录页：大标题 + 时间线式目录，更适合答辩章节切换
                addBackdrop(slide);
                slide.addText(section.title, {
                    x: 0.9, y: 1.7, w: 11.1, h: 0.9,
                    fontSize: 34, bold: true, color: theme.primary, fontFace: FONT
                });
                const bullets = section.bullets && section.bullets.length ? section.bullets : [section.visual].filter(Boolean);
                if (bullets && bullets.length) {
                    const baseY = 3.0;
                    const stepX = 2.6;
                    bullets.slice(0, 5).forEach((b, idx) => {
                        const x = 1.1 + idx * stepX;
                        slide.addShape(pptx.ShapeType.ellipse, {
                            x, y: baseY, w: 0.5, h: 0.5,
                            fill: { color: theme.accent }, line: { color: theme.accent }
                        });
                        slide.addText(String(idx + 1), {
                            x: x, y: baseY, w: 0.5, h: 0.5,
                            align: 'center', valign: 'mid', fontSize: 12, bold: true, color: theme.primary, fontFace: FONT
                        });
                        slide.addText(b, {
                            x: x - 0.4, y: baseY + 0.7, w: 1.3, h: 0.9,
                            align: 'center', fontSize: 12, color: theme.secondary, fontFace: FONT
                        });
                        if (idx < bullets.length - 1) {
                            slide.addShape(pptx.ShapeType.line, {
                                x: x + 0.5, y: baseY + 0.25, w: stepX - 0.5, h: 0,
                                line: { color: theme.secondary, pt: 1, transparency: 30 }
                            });
                        }
                    });
                }
            } else if (section._forceImage || layout.includes('图文左') || layout.includes('图文右')) {
                const imageLeft = layout.includes('图文左') || section._forceImage;
                const boxX = imageLeft ? 0.9 : 7.2;
                const txtX = imageLeft ? 7.3 : 1.0;
                addSlideWithImage(
                    slide,
                    section,
                    imageData,
                    { x: boxX, y: 2.02, w: section._forceImage ? 11.6 : 5.35, h: section._forceImage ? 3.7 : 3.35 },
                    section._forceImage ? { x: 1.0, y: 5.9, w: 11.3, h: 0.55 } : { x: boxX + 0.5, y: 5.0, w: 4.35, h: 0.95 },
                    section._forceImage ? (section.bullets.join("；") || section.visual || "建议放置高质量主视觉图片") : (section.visual || "建议放置场景图 / 产品图 / 数据图"),
                    'center',
                    section._forceImage ? 14 : 12
                );
                if (!section._forceImage) {
                    addBulletList(slide, section.bullets, txtX, 2.1, 5.1, 4.15, 15);
                }
            } else if (layout.includes('2x2')) {
                addHeader(slide, section.title, section);
                section.bullets.slice(0, 4).forEach((item, idx) => {
                    const col = idx % 2;
                    const row = Math.floor(idx / 2);
                    const x = 0.92 + col * 6.0;
                    const y = 2.0 + row * 2.2;
                    slide.addShape(pptx.ShapeType.roundRect, {
                        x, y, w: 5.5, h: 1.95, rectRadius: 0.08,
                        fill: { color: theme.secondary, transparency: 84 }, line: { color: theme.accent, pt: 1 }
                    });
                    slide.addText(item, { x: x + 0.28, y: y + 0.38, w: 4.95, h: 1.25, fontSize: 14, color: theme.secondary, fontFace: FONT });
                });
            } else if (layout.includes('时间线')) {
                // 答辩/路演常用时间线页：水平时间轴 + 节点说明
                addHeader(slide, section.title, section);
                const bullets = section.bullets || [];
                const count = Math.max(2, Math.min(5, bullets.length || 3));
                const usable = bullets.slice(0, count);
                const startX = 1.0;
                const endX = 12.0;
                const baseY = 3.1;
                slide.addShape(pptx.ShapeType.line, {
                    x: startX, y: baseY + 0.25, w: endX - startX, h: 0,
                    line: { color: theme.secondary, pt: 1.2, transparency: 15 }
                });
                usable.forEach((text, idx) => {
                    const t = idx / Math.max(1, count - 1);
                    const x = startX + t * (endX - startX);
                    slide.addShape(pptx.ShapeType.ellipse, {
                        x: x - 0.22, y: baseY, w: 0.44, h: 0.44,
                        fill: { color: theme.accent }, line: { color: theme.accent }
                    });
                    slide.addText(String(idx + 1), {
                        x: x - 0.22, y: baseY, w: 0.44, h: 0.44,
                        align: 'center', valign: 'mid', fontSize: 11, bold: true, color: theme.primary, fontFace: FONT
                    });
                    slide.addText(text, {
                        x: x - 1.4, y: baseY + 0.7, w: 2.8, h: 1.5,
                        align: 'center', fontSize: 12, color: theme.secondary, fontFace: FONT
                    });
                });
            } else if (layout.includes('数据重点')) {
                // 数据重点页：中间大数字 + 周围解释
                addHeader(slide, section.title, section);
                const bullets = section.bullets || [];
                const main = bullets[0] || section.visual || "关键数据";
                slide.addShape(pptx.ShapeType.roundRect, {
                    x: 4.3, y: 2.4, w: 5.3, h: 2.1, rectRadius: 0.2,
                    fill: { color: theme.secondary, transparency: 80 }, line: { color: theme.accent, pt: 1.4 }
                });
                slide.addText(main, {
                    x: 4.5, y: 2.65, w: 4.9, h: 1.7,
                    align: 'center', valign: 'mid',
                    fontSize: 26, bold: true, color: theme.primary, fontFace: FONT
                });
                const rest = bullets.slice(1);
                const cols = 2;
                const colW = 5.4;
                const baseY = 4.9;
                rest.slice(0, 4).forEach((t, idx) => {
                    const col = idx % cols;
                    const row = Math.floor(idx / cols);
                    const x = 1.0 + col * (colW + 1.0);
                    const y = baseY + row * 1.2;
                    slide.addText([
                        { text: "• ", options: { color: theme.accent, fontSize: 14, fontFace: FONT } },
                        { text: t, options: { color: theme.secondary, fontSize: 13, fontFace: FONT } }
                    ], { x, y, w: colW, h: 0.9, valign: 'top' });
                });
            } else if (layout.includes('对比结论')) {
                // 对比结论页：左右两列卡片式对比
                addHeader(slide, section.title, section);
                const bullets = section.bullets || [];
                const leftItems = bullets.filter((_, idx) => idx % 2 === 0);
                const rightItems = bullets.filter((_, idx) => idx % 2 === 1);
                const card = (x, title, items) => {
                    slide.addShape(pptx.ShapeType.roundRect, {
                        x, y: 2.0, w: 5.8, h: 4.0, rectRadius: 0.12,
                        fill: { color: theme.secondary, transparency: 86 }, line: { color: theme.accent, pt: 1.2 }
                    });
                    slide.addText(title, {
                        x: x + 0.4, y: 2.2, w: 5.0, h: 0.5,
                        fontSize: 16, bold: true, color: theme.primary, fontFace: FONT
                    });
                    (items || []).slice(0, 4).forEach((t, idx) => {
                        slide.addText([
                            { text: "▸ ", options: { color: theme.accent, fontSize: 13, fontFace: FONT } },
                            { text: t, options: { color: theme.secondary, fontSize: 13, fontFace: FONT } }
                        ], { x: x + 0.5, y: 2.8 + idx * 0.8, w: 4.9, h: 0.7, valign: 'top' });
                    });
                };
                card(0.9, "方案 A / 现状", leftItems);
                card(7.0, "方案 B / 目标", rightItems);
            } else if (layout.includes('紧凑要点')) {
                addHeader(slide, section.title, section);
                slide.addShape(pptx.ShapeType.roundRect, {
                    x: 0.88, y: 2.0, w: 7.7, h: 4.3, rectRadius: 0.08,
                    fill: { color: theme.secondary, transparency: 86 }, line: { color: theme.accent, pt: 1 }
                });
                slide.addShape(pptx.ShapeType.roundRect, {
                    x: 8.78, y: 2.0, w: 3.75, h: 4.3, rectRadius: 0.08,
                    fill: { color: theme.secondary, transparency: 90 }, line: { color: theme.accent, pt: 1 }
                });
                addBulletList(slide, section.bullets, 1.14, 2.34, 7.2, 3.72, 16);
                slide.addText(section.visual || "关键洞察", {
                    x: 9.02, y: 2.45, w: 3.3, h: 0.6, fontSize: 14, bold: true, color: theme.accent, fontFace: FONT
                });
            } else if (layout.includes('总结收束') || pageType.includes('总结') || pageType.includes('q&a')) {
                // 总结 / Q&A：居中大标题 + 单栏要点，更克制的收束页
                addBackdrop(slide);
                slide.addText(section.title, {
                    x: 0.9, y: 1.9, w: 11.1, h: 0.9,
                    fontSize: 32, bold: true, color: theme.primary, fontFace: FONT, align: 'center'
                });
                const bullets = section.bullets || [];
                if (bullets.length) {
                    addBulletList(slide, bullets, 2.0, 3.0, 9.0, 3.0, 17);
                } else if (layout.includes('q&a') || pageType.includes('q&a')) {
                    slide.addText("Q & A", {
                        x: 0.9, y: 3.0, w: 11.1, h: 1.5,
                        fontSize: 60, bold: true, color: theme.accent, fontFace: FONT, align: 'center'
                    });
                }
            } else {
                addHeader(slide, section.title, section);
                const mid = Math.ceil(section.bullets.length / 2);
                addBulletList(slide, section.bullets.slice(0, mid), 1.0, 2.1, 5.45, 4.1, 15);
                addBulletList(slide, section.bullets.slice(mid), 6.9, 2.1, 5.2, 4.1, 15);
            }

            addVisualBadge(slide, section);
            addMiniChart(slide, section);
            addBigNumberKpi(slide, section);

            if (section.note) {
                slide.addText(section.note, {
                    x: 0.92, y: 6.48, w: 9.8, h: 0.3,
                    fontSize: 8.5, color: theme.secondary, opacity: 0.7, fontFace: FONT
                });
            }
            addFooter(slide, i, sections.length);
        }

        const fileName = `${String(topic || "AI高质感PPT").replace(/[^\w\u4e00-\u9fa5-]/g, "_")}_${Date.now()}.pptx`;
        const buf = await pptx.write({ outputType: 'nodebuffer' });
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
        res.setHeader('Content-Disposition', `attachment; filename=\"${encodeURIComponent(fileName)}\"`);
        return res.send(buf);
    } catch (error) {
        console.error("export-ppt 接口报错:", error);
        return res.status(500).json({ error: "PPT 导出失败，请稍后重试" });
    }
});

app.get('/stock-image', async (req, res) => {
    const q = String(req.query.q || "").trim();
    const keyword = normalizeImageQuery(q);
    const seed = encodeURIComponent((q || keyword).slice(0, 60));

    const candidates = [
        `https://loremflickr.com/1600/900/${encodeURIComponent(keyword)}`,
        `https://source.unsplash.com/1600x900/?${encodeURIComponent(keyword)}`,
        `https://picsum.photos/seed/${seed}/1600/900`
    ];

    for (const url of candidates) {
        try {
            const controller = new AbortController();
            const timer = setTimeout(() => controller.abort(), 8000);
            const resp = await fetch(url, { signal: controller.signal, redirect: 'follow' });
            clearTimeout(timer);
            const type = resp.headers.get('content-type') || '';
            if (!resp.ok || !type.startsWith('image/')) continue;
            const arrBuf = await resp.arrayBuffer();
            res.setHeader('Content-Type', type);
            res.setHeader('Cache-Control', 'public, max-age=3600');
            return res.send(Buffer.from(arrBuf));
        } catch (_) {
            continue;
        }
    }

    return res.status(502).json({ error: "暂时无法获取配图" });
});

// 3. 启动服务器
app.listen(3000, () => {
    console.log('🚀 服务器运行在: http://localhost:3000');
});

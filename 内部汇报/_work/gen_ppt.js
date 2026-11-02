const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.author = "FDE";
pres.title = "商业模式五层级·五问汇报";

const W = 13.33, H = 7.5, M = 0.55;
const F = "微软雅黑";
const INK = "10231C", PRIMARY = "1D5C45";
const TINT = "E9F2ED", TINT2 = "CFE3D8";
const ACC = "E8A33D", ACC_D = "A66A12";
const TEXT = "22332C", MUTED = "6E837A", FAINT = "9DBBAC", WHITE = "FFFFFF";
const AMB_T = "FBF2E2";
const sh = () => ({ type: "outer", color: "1A2E28", blur: 7, offset: 2, angle: 60, opacity: 0.14 });
const S = pres.ShapeType;

function pageTitle(s, q, title, page) {
  if (q) {
    s.addText(q, { x: M, y: 0.4, w: 1.2, h: 0.78, fontFace: F, fontSize: 40, bold: true, color: ACC, margin: 0, valign: "middle", align: "left" });
    s.addText(title, { x: M + 1.28, y: 0.4, w: 10.5, h: 0.78, fontFace: F, fontSize: 30, bold: true, color: TEXT, margin: 0, valign: "middle", align: "left" });
  } else {
    s.addText(title, { x: M, y: 0.4, w: 11.5, h: 0.78, fontFace: F, fontSize: 30, bold: true, color: TEXT, margin: 0, valign: "middle", align: "left" });
  }
  if (page) s.addText(page, { x: W - 1.45, y: 7.05, w: 0.9, h: 0.32, fontFace: F, fontSize: 12, color: MUTED, margin: 0, align: "right" });
}

function card(s, x, y, w, h, fill, line) {
  const o = { x, y, w, h, fill: { color: fill }, rectRadius: 0.07, shadow: sh() };
  if (line) o.line = { color: line, width: 1 }; else o.line = { type: "none" };
  s.addShape(S.roundRect, o);
}

/* ============ S1 封面 ============ */
{
  const s = pres.addSlide();
  s.background = { color: INK };
  // 五级台阶 motif（右下）
  const hs = [0.5, 0.92, 1.34, 1.76, 2.2];
  const cs = ["1F4E3C", "2E6B52", "3E7C63", "66A88C", ACC];
  for (let i = 0; i < 5; i++) {
    s.addShape(S.rect, { x: 8.95 + i * 0.86, y: 6.78 - hs[i], w: 0.72, h: hs[i], fill: { color: cs[i] }, line: { type: "none" } });
  }
  s.addText("课 程 学 习 汇 报 · 内 部 资 料", { x: M, y: 1.25, w: 7, h: 0.4, fontFace: F, fontSize: 13, color: FAINT, margin: 0, charSpacing: 3 });
  s.addText([
    { text: "从五个问题，", options: { breakLine: true } },
    { text: "看我们要成为什么样的公司" },
  ], { x: M, y: 1.75, w: 9.6, h: 2.3, fontFace: F, fontSize: 46, bold: true, color: WHITE, margin: 0, lineSpacingMultiple: 1.15, align: "left", valign: "top" });
  s.addText("商业模式五层级 ——《创业者认知升级与商业模式分析》学习理解", { x: M, y: 4.25, w: 9.6, h: 0.5, fontFace: F, fontSize: 16, color: "CFE0D8", margin: 0 });
  s.addText("依据：混沌创商院二模块课程纪要（飞书妙记 · 2026-10-23）　　汇报人：＿＿＿＿　　2026-10-25", { x: M, y: 6.62, w: 8.2, h: 0.4, fontFace: F, fontSize: 12, color: "7E998D", margin: 0 });
  s.addNotes("【讲稿】开场：各位好。我用十分钟汇报一下贾光校长这次课的学习。老板留了五个问题，我按顺序讲我自己的理解，最后把五个问题合成一张图。先亮结论：这五个问题其实是同一个问题——我们要成为一家什么样的公司。\n【提示】语速放慢，说完结论停一秒再翻页。右下角的台阶图形后面还会出现，是这堂课的核心模型。");
}

/* ============ S2 一页看懂这门课 ============ */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  pageTitle(s, null, "一页看懂这门课", "02 / 10");
  s.addText("商业模式的竞争，是最高维度的竞争", { x: M, y: 1.3, w: 12.2, h: 0.68, fontFace: F, fontSize: 33, bold: true, color: PRIMARY, margin: 0 });
  s.addText("产品竞争是低维度的——任何产品（包括高科技产品）终将被复刻", { x: M, y: 2.0, w: 12.2, h: 0.42, fontFace: F, fontSize: 15, color: MUTED, margin: 0 });

  card(s, M, 2.62, 6.0, 1.85, TINT);
  s.addText("赚钱", { x: M + 0.32, y: 2.85, w: 2, h: 0.5, fontFace: F, fontSize: 21, bold: true, color: PRIMARY, margin: 0 });
  s.addText("当期现金流。层级越低越容易赚钱——但天花板也越低，难对接资本市场。", { x: M + 0.32, y: 3.42, w: 5.4, h: 0.85, fontFace: F, fontSize: 13, color: TEXT, margin: 0, lineSpacingMultiple: 1.15 });
  card(s, 6.78, 2.62, 6.0, 1.85, INK);
  s.addText("值钱", { x: 7.1, y: 2.85, w: 2, h: 0.5, fontFace: F, fontSize: 21, bold: true, color: ACC, margin: 0 });
  s.addText("资本为“所有参与者带来的未来收益”定的价。层级越高越值钱，甚至可以不赚钱。", { x: 7.1, y: 3.42, w: 5.4, h: 0.85, fontFace: F, fontSize: 13, color: WHITE, margin: 0, lineSpacingMultiple: 1.15 });

  const items = [
    ["世界观", "从确定性到不确定性：“因上努力，果上概率”"],
    ["盈利逻辑", "信息差红利结束，认知红利开始"],
    ["事业定位", "区分生意与事业，事业要延迟满足"],
  ];
  items.forEach((it, i) => {
    const x = M + i * 4.18;
    s.addShape(S.rect, { x, y: 4.95, w: 0.18, h: 0.18, fill: { color: ACC }, line: { type: "none" } });
    s.addText(it[0], { x: x + 0.32, y: 4.85, w: 3.6, h: 0.4, fontFace: F, fontSize: 15, bold: true, color: PRIMARY, margin: 0 });
    s.addText(it[1], { x: x + 0.32, y: 5.3, w: 3.55, h: 0.75, fontFace: F, fontSize: 12.5, color: TEXT, margin: 0, lineSpacingMultiple: 1.12 });
  });
  s.addText("来源：飞书妙记《创业者认知升级与商业模式分析》智能纪要（2026-10-23）", { x: M, y: 6.55, w: 10, h: 0.35, fontFace: F, fontSize: 12, color: MUTED, margin: 0 });
  s.addNotes("【讲稿】先用一分钟说这门课在讲什么。核心就一句话：商业模式的竞争是最高维度的竞争，产品竞争是低维度的，任何产品包括高科技产品最终都会被复刻。他给了一把尺子，叫五个层级，下一页展开。这一页先记两个词：赚钱和值钱。赚钱是当期现金流；值钱是资本市场为“未来所有参与者能带来的收益”付的价格。层级越低越赚钱但不值钱，层级越高越值钱，甚至可以不赚钱——这句话是整堂课的主线，也是后面五个问题的底色。\n【提示】指一下两张对比卡片；“甚至可以不赚钱”停半拍，这是反直觉的点。");
}

/* ============ S3 Q1 ============ */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  pageTitle(s, "Q1", "为什么需要了解这个问题", "03 / 10");

  card(s, M, 1.5, 4.4, 1.2, TINT);
  s.addText("信息差时代", { x: M + 0.28, y: 1.66, w: 3.8, h: 0.42, fontFace: F, fontSize: 17, bold: true, color: PRIMARY, margin: 0 });
  s.addText("靠信息不对称赚差价", { x: M + 0.28, y: 2.12, w: 3.8, h: 0.38, fontFace: F, fontSize: 12.5, color: MUTED, margin: 0 });
  s.addShape(S.rightArrow, { x: 5.2, y: 1.95, w: 2.5, h: 0.34, fill: { color: ACC }, line: { type: "none" } });
  s.addText("AI 信息平权", { x: 5.2, y: 1.52, w: 2.5, h: 0.36, fontFace: F, fontSize: 12.5, bold: true, color: ACC_D, margin: 0, align: "center" });
  card(s, 8.0, 1.5, 4.78, 1.2, INK);
  s.addText("认知差时代", { x: 8.28, y: 1.66, w: 4.2, h: 0.42, fontFace: F, fontSize: 17, bold: true, color: ACC, margin: 0 });
  s.addText("同样的信息，看得更深的人赢", { x: 8.28, y: 2.12, w: 4.2, h: 0.38, fontFace: F, fontSize: 12.5, color: WHITE, margin: 0 });

  const rows = [
    ["1", "时代变了", "AI 让信息全面平权，靠信息差赚差价越来越走不通，未来拼的是认知深度。"],
    ["2", "我们的位置", "FDE 服务天然长在第 1 层——卖的是工程师的时间，时间是天花板。"],
    ["3", "问题的分量", "层级决定“我们凭什么值钱”，决定增长方式是加人，还是降边际成本。"],
  ];
  rows.forEach((r, i) => {
    const y = 3.1 + i * 1.28;
    s.addShape(S.roundRect, { x: M, y, w: 0.62, h: 0.62, fill: { color: PRIMARY }, line: { type: "none" }, rectRadius: 0.06 });
    s.addText(r[0], { x: M, y, w: 0.62, h: 0.62, fontFace: F, fontSize: 22, bold: true, color: WHITE, margin: 0, align: "center", valign: "middle" });
    s.addText(r[1], { x: M + 0.92, y: y - 0.02, w: 3.2, h: 0.5, fontFace: F, fontSize: 16.5, bold: true, color: TEXT, margin: 0 });
    s.addText(r[2], { x: M + 0.92, y: y + 0.46, w: 11.2, h: 0.55, fontFace: F, fontSize: 13.5, color: MUTED, margin: 0 });
    if (i < 2) s.addShape(S.line, { x: M + 0.92, y: y + 1.12, w: 11.2, h: 0, line: { color: "E4EDE8", width: 0.75 } });
  });
  s.addNotes("【讲稿】第一个问题，为什么需要了解这个问题。我理解有三层。第一层是时代：AI 把信息拉平了，靠信息差赚差价的时代结束，接下来拼的是同样一段信息看得比别人深，也就是认知红利。第二层是我们自己的位置：我们是 FDE 服务模式，天然长在第一层，卖的是工程师的时间，时间是天花板——不了解层级问题，就会一直用“接更多项目、加更多人”的方式增长。第三层是这个问题本身的分量：它其实在问“我们凭什么值钱”，答案决定我们的增长方式是加人，还是降边际成本。\n【提示】三条按“时代—位置—分量”的顺序讲，第2条最贴近我们自己，可以多停一下。");
}

/* ============ S4 Q2-a 五层级阶梯 ============ */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  pageTitle(s, "Q2", "为什么是这样的五个层级", "04 / 10");
  s.addText("台阶越高：可复制性 ↑　边际成本 ↓　从“赚钱”走向“值钱”", { x: 4.4, y: 1.32, w: 8.4, h: 0.4, fontFace: F, fontSize: 13.5, bold: true, color: ACC_D, margin: 0, align: "right" });

  const names = ["非标定制服务", "标准化产品", "产品＋服务", "双边平台", "生态"];
  const feats = [
    "时间是天花板\n赚钱但估值低",
    "可跨区域复制\n但终陷价格战",
    "拉高终身价值\n建私域、抗比价",
    "交易抽成“雁过拔毛”\n需 12 字准则成立",
    "像水和空气\n离不开",
  ];
  const typs = ["咨询 · 律所", "标准品公司", "产品化 SaaS", "网约车 · 家装平台", "微信 · 苹果"];
  const fills = ["D9E8E0", "BCD8CA", "8FC0A9", "4E8F71", "1D5C45"];
  const bottoms = 6.35, heights = [1.95, 2.44, 2.93, 3.46, 3.98];
  for (let i = 0; i < 5; i++) {
    const x = 0.7 + i * 2.38, hgt = heights[i], y = bottoms - hgt;
    s.addShape(S.rect, { x, y, w: 2.28, h: hgt, fill: { color: fills[i] }, line: { type: "none" }, shadow: sh() });
    const dark = i >= 3;
    s.addText(`第${i + 1}层`, { x: x + 0.14, y: y + 0.1, w: 2.0, h: 0.3, fontFace: F, fontSize: 12, bold: true, color: dark ? (i === 4 ? ACC : "D8EDE2") : "4E7A66", margin: 0 });
    s.addText(names[i], { x: x + 0.14, y: y + 0.4, w: 2.0, h: 0.4, fontFace: F, fontSize: 16.5, bold: true, color: dark ? WHITE : TEXT, margin: 0 });
    const lines = feats[i].split("\n").map((t, k, arr) => ({ text: t, options: { breakLine: k < arr.length - 1 } }));
    s.addText(lines, { x: x + 0.14, y: y + 0.84, w: 2.0, h: 0.52, fontFace: F, fontSize: 12, color: dark ? "EAF4EF" : "3C544A", margin: 0, lineSpacingMultiple: 1.1 });
    s.addText(typs[i], { x: x + 0.14, y: y + 1.42, w: 2.0, h: 0.32, fontFace: F, fontSize: 12, color: dark ? "CFE5DB" : MUTED, margin: 0 });
  }
  s.addText("越低越“赚钱”，但天花板低、难对接资本市场", { x: 0.7, y: 6.5, w: 7.5, h: 0.36, fontFace: F, fontSize: 12.5, color: MUTED, margin: 0 });
  s.addText("课程注：另有第 6 层“垄断”（国家电网、高铁）——官方主导，民企只能在小的领域做阶段性垄断", { x: 0.7, y: 6.95, w: 10.6, h: 0.34, fontFace: F, fontSize: 12, color: MUTED, margin: 0 });
  s.addNotes("【讲稿】第二个问题前半：为什么是这五个层级。从低到高过一遍：第一层非标定制服务，咨询、律所，时间是天花板，赚钱但估值低；第二层标准化产品，可以跨区域复制，但最终陷入同质化价格战；第三层产品加服务，用可复制的服务拉高用户终身价值、建私域，抵御比价；第四层双边平台，靠交易抽成，成立要满足十二个字：上游过剩、下游分散、动态匹配；第五层生态，像微信和苹果，所有参与者都在为它创造收益。课程里还有第六层“垄断”，国家电网、高铁那种，民企碰不到，知道就行。这五层共同的刻度就一句话：可复制性越来越强、边际成本越来越低，公司从赚钱走向值钱。\n【提示】手指沿着台阶从左下往右上走；“上游过剩、下游分散、动态匹配”这12个字后面第6页还会用回来。");
}

/* ============ S5 Q2-b 我们在哪里 ============ */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  pageTitle(s, "Q2", "我们现在在哪里", "05 / 10");

  const mh = [0.85, 1.3, 1.75, 2.2, 2.65];
  const mf = [ACC, TINT2, TINT2, "DEEAE4", "DEEAE4"];
  for (let i = 0; i < 5; i++) {
    const x = 0.75 + i * 1.06, y = 6.05 - mh[i];
    s.addShape(S.rect, { x, y, w: 0.98, h: mh[i], fill: { color: mf[i] }, line: { type: "none" } });
    s.addText(String(i + 1), { x, y: y + 0.06, w: 0.98, h: 0.3, fontFace: F, fontSize: 12.5, bold: true, color: i === 0 ? WHITE : "7E998D", margin: 0, align: "center" });
  }
  s.addText("我们在这", { x: 0.75, y: 4.78, w: 1.3, h: 0.36, fontFace: F, fontSize: 13, bold: true, color: ACC_D, margin: 0 });
  s.addShape(S.line, { x: 1.4, y: 4.0, w: 1.8, h: 0.6, flipV: true, line: { color: ACC, width: 2.25, endArrowType: "triangle" } });
  s.addText("向 2–3 层爬", { x: 1.6, y: 3.52, w: 1.7, h: 0.36, fontFace: F, fontSize: 12.5, bold: true, color: ACC_D, margin: 0 });
  s.addText("主体在第 1 层，正向 2–3 层过渡", { x: 0.55, y: 6.35, w: 5.4, h: 0.38, fontFace: F, fontSize: 12.5, color: MUTED, margin: 0, align: "center" });

  const ev = [
    ["已经在做 ①　知识库 / SOP / 复盘台账", "把 FDE 方法论标准化——这是从第 1 层往第 2 层走的动作。", TINT, PRIMARY],
    ["已经在做 ②　闭环跑在我们的系统上", "“客户停付费，闭环就断”——这已经是第 3→4 层的思维。", TINT, PRIMARY],
    ["自检判据：新接客户的边际成本，是否递减", "每接一个客户就加一队人＝还在 1 层；接第 5 个客户只要第 1 个的 1/3 人力＝在爬层。", AMB_T, ACC_D],
  ];
  ev.forEach((e, i) => {
    const y = 1.5 + i * 1.72;
    card(s, 6.35, y, 6.43, 1.55, e[2]);
    s.addText(e[0], { x: 6.67, y: y + 0.16, w: 5.85, h: 0.42, fontFace: F, fontSize: 15, bold: true, color: e[3], margin: 0 });
    s.addText(e[1], { x: 6.67, y: y + 0.62, w: 5.8, h: 0.8, fontFace: F, fontSize: 13, color: TEXT, margin: 0, lineSpacingMultiple: 1.12 });
  });
  s.addNotes("【讲稿】后半：我们在哪里。诚实的回答：主体在第一层，但已经在往二、三层爬，有两个证据。一是我们正在做的知识库、SOP、复盘台账，本质是把方法论标准化，这是从一层往二层的动作；二是我们自己方法论里的那句话——“闭环必须跑在我们的系统上，客户停付费闭环就断”——这已经不是一层的思维，这是三层往四层的思维，服务和系统有了粘性。再给自己一个自检判据：新接一个客户的边际成本是否在递减。每接一个客户就加一队人，那还是一层；接第五个客户只要第一个客户三分之一的人力，才叫在爬层。\n【提示】这页是五个问题里最需要“自己的话”的一页，语气要笃定，但结论要诚实——不夸大我们现在的高度。");
}

/* ============ S6 Q3 ============ */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  pageTitle(s, "Q3", "我们能去到哪里", "06 / 10");

  card(s, M, 1.45, 12.23, 1.32, TINT);
  s.addText("愿景：成为中小电商品牌的经营基础设施", { x: M + 0.32, y: 1.6, w: 8.8, h: 0.48, fontFace: F, fontSize: 19, bold: true, color: PRIMARY, margin: 0 });
  s.addText("形态：品类经营生态（第 5 层）＝ 数据基座 × 开发者定制应用", { x: M + 0.32, y: 2.12, w: 8.8, h: 0.45, fontFace: F, fontSize: 13.5, color: TEXT, margin: 0 });
  s.addText("L5", { x: 10.6, y: 1.52, w: 1.9, h: 0.85, fontFace: F, fontSize: 44, bold: true, color: ACC, margin: 0, align: "right" });
  s.addText("生态", { x: 10.6, y: 2.32, w: 1.9, h: 0.4, fontFace: F, fontSize: 13, bold: true, color: MUTED, margin: 0, align: "right" });

  const cw = [
    ["冷水 ①　数据基座是终点，不是起点", "开发者跟着用户走——先把客户闭环跑在我们系统上，飞轮转起来，才轮到谈开放。"],
    ["冷水 ②　我们独有的只有一类数据：行动-结果对", "谁、在哪个环节、依据什么判断、做了什么动作、结果如何——平台看不见，ERP 也看不见。"],
  ];
  cw.forEach((c, i) => {
    const x = M + i * 6.23;
    card(s, x, 3.0, 6.0, 1.75, WHITE, TINT2);
    s.addText(c[0], { x: x + 0.3, y: 3.18, w: 5.5, h: 0.44, fontFace: F, fontSize: 15, bold: true, color: TEXT, margin: 0 });
    s.addText(c[1], { x: x + 0.3, y: 3.66, w: 5.42, h: 0.95, fontFace: F, fontSize: 13, color: MUTED, margin: 0, lineSpacingMultiple: 1.15 });
  });

  s.addText("平台 12 字检验：上游过剩 ✓　下游分散 ✓　动态匹配＝机会——缺的是双边都不在场", { x: M, y: 4.95, w: 12.2, h: 0.4, fontFace: F, fontSize: 13.5, bold: true, color: PRIMARY, margin: 0 });

  const steps = ["① 闭环跑上系统", "② 多客户数据飞轮", "③ 开发者主动接入＝生态门口"];
  const sf = [TINT2, TINT2, AMB_T];
  const sc = [TEXT, TEXT, ACC_D];
  steps.forEach((t, i) => {
    const x = M + i * 4.3;
    s.addShape(S.roundRect, { x, y: 5.55, w: 3.62, h: 0.72, fill: { color: sf[i] }, line: { type: "none" }, rectRadius: 0.08 });
    s.addText(t, { x, y: 5.55, w: 3.62, h: 0.72, fontFace: F, fontSize: 13.5, bold: true, color: sc[i], margin: 0, align: "center", valign: "middle" });
    if (i < 2) s.addShape(S.rightArrow, { x: x + 3.72, y: 5.76, w: 0.48, h: 0.3, fill: { color: ACC }, line: { type: "none" } });
  });
  s.addNotes("【讲稿】第三问，能去到哪里。我的答案：方向是品类经营生态，第五层——数据基座加开发者定制应用，成为中小电商品牌的经营基础设施。但这里我给自己泼了两盆冷水。第一盆：数据基座是终点不是起点，开发者永远跟着用户走，得先把客户的闭环跑在我们系统上，第三家、第五家客户上来，数据飞轮转起来，才轮到开放。第二盆：“数据”这个词要拆开——源头数据在平台手里，巨量引擎有投放数据、生意参谋有成交数据，我们能调的都是接口里的二手数据；我们真正独有的只有一类，叫行动-结果对：谁在哪个环节、依据什么判断、做了什么动作、结果如何。这是经营决策过程数据，平台看不见，ERP 也看不见。用十二字准则检验：上游过剩成立、下游分散成立，缺的是双边都不在场。所以里程碑是三步：闭环上系统、多客户数据飞轮、开发者主动来接入——到那才叫生态门口。\n【提示】“行动-结果对”是这页的关键词，说慢一点；两盆冷水体现的是我想过这个愿景的漏洞，不是退缩。");
}

/* ============ S7 Q4 ============ */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  pageTitle(s, "Q4", "怎么去到那里：三步走", "07 / 10");

  const st = [
    ["1", "品类做深", "女士内衣：SPU 主键、投放节奏、复盘台账，沉淀成结构化资产——做到品类里提起“经营系统”就想起我们。"],
    ["2", "闭环系统化", "以代运营为形态、以系统为资产：服务费是现金流，系统里的闭环与数据才是资产。"],
    ["3", "横向复制", "进相邻品类时，带走的是模板和数据，不是人——边际成本才会递减。"],
  ];
  st.forEach((t, i) => {
    const x = M + i * 4.15;
    card(s, x, 1.45, 3.95, 2.25, TINT);
    s.addShape(S.roundRect, { x: x + 0.28, y: 1.7, w: 0.56, h: 0.56, fill: { color: PRIMARY }, line: { type: "none" }, rectRadius: 0.06 });
    s.addText(t[0], { x: x + 0.28, y: 1.7, w: 0.56, h: 0.56, fontFace: F, fontSize: 22, bold: true, color: WHITE, margin: 0, align: "center", valign: "middle" });
    s.addText(t[1], { x: x + 1.0, y: 1.76, w: 2.8, h: 0.48, fontFace: F, fontSize: 17, bold: true, color: TEXT, margin: 0 });
    s.addText(t[2], { x: x + 0.28, y: 2.42, w: 3.42, h: 1.15, fontFace: F, fontSize: 12.5, color: TEXT, margin: 0, lineSpacingMultiple: 1.15 });
  });

  s.addText("课程里的四个升级案例", { x: M, y: 3.92, w: 6, h: 0.38, fontFace: F, fontSize: 13.5, bold: true, color: MUTED, margin: 0 });
  const cases = [
    ["律所", "定制 200–300 万 → 3.8 万标准品，快速起量，拿到天使投资。"],
    ["贝思特肉松", "卖产品 3 亿触顶 → 转做烘焙店经营方案，10 亿且未降价。"],
    ["购美", "代运营不当乙方：数据中台预测爆款，反向签下头部品牌。"],
    ["抖音翡翠", "单月卖货 4000 万 → 同行涌入，3 个月转型行业代运营平台。"],
  ];
  cases.forEach((c, i) => {
    const x = M + i * 3.08;
    card(s, x, 4.32, 2.93, 1.78, WHITE, TINT2);
    s.addText(c[0], { x: x + 0.24, y: 4.5, w: 2.5, h: 0.4, fontFace: F, fontSize: 14, bold: true, color: PRIMARY, margin: 0 });
    s.addText(c[1], { x: x + 0.24, y: 4.94, w: 2.48, h: 1.05, fontFace: F, fontSize: 12, color: TEXT, margin: 0, lineSpacingMultiple: 1.12 });
  });

  s.addShape(S.rect, { x: 0, y: 6.32, w: W, h: 0.62, fill: { color: AMB_T }, line: { type: "none" } });
  s.addText("避坑：定制软件＝第 1 层披着第 2 层的外衣——既没有服务粘性，也没有产品的边际成本递减。", { x: M, y: 6.32, w: 12.2, h: 0.62, fontFace: F, fontSize: 13, bold: true, color: ACC_D, margin: 0, valign: "middle" });
  s.addNotes("【讲稿】第四问，怎么去。三步：第一步品类做深，就在女士内衣，把 SPU 主键、投放节奏、复盘台账变成结构化资产，做到这个品类里品牌提起经营系统就想起我们；第二步闭环系统化，定位是以代运营为形态、以系统为资产——服务费是现金流，系统里的闭环和数据才是资产；第三步横向复制，进相邻品类时带走的是模板和数据，不是人，这样边际成本才递减。课程里有四个现成案例支撑这条路：律所把两三百万的定制服务做成 3.8 万的标准品，拿到天使投资；贝思特肉松从卖产品 3 亿到做烘焙店经营方案 10 亿；购美从代运营做成用数据中台反向签头部品牌；抖音卖翡翠的学员三个月转成行业代运营平台。要避的坑也明确：别掉进定制软件，那是第一层披着第二层的外衣，两头不占。\n【提示】三个步骤之间是递进关系，讲完每步可以看一眼听众；四个案例快速带过，每一个都是“同一个动作：往上一层”。");
}

/* ============ S8 Q5-a ============ */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  pageTitle(s, "Q5", "自产、自销、自营，是好公司吗？", "08 / 10");

  s.addText("不是好公司的判断标准。", { x: M, y: 1.38, w: 12.2, h: 0.62, fontFace: F, fontSize: 28, bold: true, color: TEXT, margin: 0 });
  card(s, M, 2.12, 12.23, 1.18, INK);
  s.addText([
    { text: "“推崇开放生态，反对全链条自营——全链条自营让公司资源受限、愈发脆弱。”", options: { breakLine: true } },
    { text: "—— 课程纪要 · 01:19:09 章节", options: { color: FAINT, fontSize: 12 } },
  ], { x: M + 0.35, y: 2.12, w: 11.5, h: 1.18, fontFace: F, fontSize: 16, bold: true, color: WHITE, margin: 0, valign: "middle", lineSpacingMultiple: 1.25 });

  const facts = [
    ["多数品牌在抖音并不盈利", "利润被平台（高阶模式）赚走——自产自销的品牌，在给平台打工。", TINT, PRIMARY],
    ["身边样本：高乐捷", "100 多人、人无限分散，连一个完整业务闭环的人都凑不齐。", TINT, ACC_D],
  ];
  facts.forEach((f, i) => {
    const x = M + i * 6.23;
    card(s, x, 3.55, 6.0, 1.62, f[2]);
    s.addText(f[0], { x: x + 0.3, y: 3.72, w: 5.5, h: 0.44, fontFace: F, fontSize: 15.5, bold: true, color: f[3], margin: 0 });
    s.addText(f[1], { x: x + 0.3, y: 4.2, w: 5.42, h: 0.8, fontFace: F, fontSize: 13, color: TEXT, margin: 0, lineSpacingMultiple: 1.12 });
  });

  s.addShape(S.rect, { x: 0, y: 5.45, w: W, h: 1.3, fill: { color: AMB_T }, line: { type: "none" } });
  s.addText([
    { text: "辨析：自营 ≠ 原罪。苹果也自研自售，但它值钱靠的是 App Store 生态——所有参与者都在为它创造收益。", options: { bold: true, breakLine: true } },
    { text: "判断标准是层级与“值钱”，不是“是否自营”；高乐捷的病根也不是自营，是决策分散＋没有系统承接高频决策。", options: {} },
  ], { x: M, y: 5.45, w: 12.2, h: 1.3, fontFace: F, fontSize: 13.5, color: ACC_D, margin: 0, valign: "middle", lineSpacingMultiple: 1.3 });
  s.addNotes("【讲稿】第五问，自产、自销、自营是不是好公司。我的回答：它不是好公司的判断标准。课程原话是推崇开放生态、反对全链条自营——全链条自营让公司资源受限、愈发脆弱。两个注脚：一个是多数品牌在抖音都不盈利，利润被平台赚走了，自产自销的品牌其实在给高阶模式打工；另一个就是我们身边的高乐捷，一百多号人，人无限分散，连一个完整业务闭环的人都凑不齐。但这里要做个辨析：自营本身不是原罪，苹果也自研自售，但苹果值钱不是因为自营，是 App Store 生态里所有参与者都在为它创造收益。所以判断一家公司好不好，看的不是“是否自营”，看的是在第几层、是否值钱。高乐捷的病根也不是自营，是决策分散——每个环节都要有人在场拍板，人当然不够用。\n【提示】先给结论再给引文，引用时放慢；高乐捷是我们自己的客户，讲的时候对事不对人。");
}

/* ============ S9 Q5-b ============ */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  pageTitle(s, "Q5", "那该放什么、不放什么", "09 / 10");

  card(s, M, 1.45, 5.95, 2.6, PRIMARY);
  s.addText("不放：不可逆的方向决策", { x: M + 0.32, y: 1.66, w: 5.3, h: 0.48, fontFace: F, fontSize: 17, bold: true, color: WHITE, margin: 0 });
  s.addText([
    { text: "品类定义　·　定价权", options: { breakLine: true } },
    { text: "品牌资产　·　融资节奏", options: { breakLine: true } },
  ], { x: M + 0.32, y: 2.28, w: 5.3, h: 0.95, fontFace: F, fontSize: 14.5, bold: true, color: "EAF4EF", margin: 0, lineSpacingMultiple: 1.45 });
  s.addText("——这是公司的本体，一丝都不能让", { x: M + 0.32, y: 3.4, w: 5.3, h: 0.4, fontFace: F, fontSize: 12.5, color: "AECBBE", margin: 0 });

  card(s, 6.83, 1.45, 5.95, 2.6, AMB_T);
  s.addText("放：高频经营决策", { x: 7.15, y: 1.66, w: 5.3, h: 0.48, fontFace: F, fontSize: 17, bold: true, color: ACC_D, margin: 0 });
  s.addText([
    { text: "投放执行　·　内容量产", options: { breakLine: true } },
    { text: "客服　·　自播日常运营", options: { breakLine: true } },
  ], { x: 7.15, y: 2.28, w: 5.3, h: 0.95, fontFace: F, fontSize: 14.5, bold: true, color: TEXT, margin: 0, lineSpacingMultiple: 1.45 });
  s.addText("外包给代理商是初级形态，最终形态是外包给系统与数据", { x: 7.15, y: 3.4, w: 5.4, h: 0.4, fontFace: F, fontSize: 12.5, color: ACC_D, margin: 0 });

  s.addText("放完之后，公司剩下三样——让出任何一样，公司就退化成一个壳：", { x: M, y: 4.32, w: 12.2, h: 0.4, fontFace: F, fontSize: 13.5, color: MUTED, margin: 0 });
  const chips = ["方向决策权", "核心资产所有权", "收益分配权"];
  chips.forEach((c, i) => {
    const x = M + i * 4.18;
    s.addShape(S.roundRect, { x, y: 4.78, w: 3.88, h: 0.72, fill: { color: INK }, line: { type: "none" }, rectRadius: 0.09, shadow: sh() });
    s.addText(c, { x, y: 4.78, w: 3.88, h: 0.72, fontFace: F, fontSize: 15.5, bold: true, color: WHITE, margin: 0, align: "center", valign: "middle" });
  });

  card(s, M, 5.72, 12.23, 1.15, WHITE, TINT2);
  s.addText("高乐捷的路线：学婷美，但要更进一步", { x: M + 0.3, y: 5.86, w: 11.6, h: 0.4, fontFace: F, fontSize: 13.5, bold: true, color: PRIMARY, margin: 0 });
  s.addText("放生产、放销售执行；但用户数据与关系必须回流自己手里——代理可以换，数据不能丢。做不到这一条的下放，叫让渡，不叫共赢。", { x: M + 0.3, y: 6.28, w: 11.6, h: 0.45, fontFace: F, fontSize: 12.5, color: TEXT, margin: 0 });
  s.addNotes("【讲稿】那该放什么、不放什么？决策要拆成两类：不可逆的方向决策不能放——品类定义、定价权、品牌资产、融资节奏，这是公司的本体；高频经营决策恰恰最该放——投不投、加不加预算、这周调什么，这类决策外包给代理商只是初级形态，最终形态是外包给系统和数据。放完之后公司剩下三样东西：方向决策权、核心资产所有权、收益分配权——这三样让出去任何一样，公司就退化成一个壳。落到高乐捷：可以学婷美的结构，放生产、放销售执行，但要比婷美更进一步——用户数据和关系必须回流自己手里，代理可以换，数据不能丢。做不到这一条的下放不叫共赢，叫让渡。\n【提示】左右两栏对照着讲，手势一左一右；“三样”逐个念；婷美那句话是这页的落点。");
}

/* ============ S10 收尾 ============ */
{
  const s = pres.addSlide();
  s.background = { color: INK };
  s.addText("一页收尾：我们的战略叙事", { x: M, y: 0.55, w: 11, h: 0.7, fontFace: F, fontSize: 30, bold: true, color: WHITE, margin: 0 });

  const nodes = [
    ["今天 · 第 1 层", "替品牌执行闭环", "1C4033", WHITE, "5E7A6E"],
    ["第 3 层", "把闭环固化成系统", "1F4E3C", WHITE, null],
    ["第 4 层", "多客户数据养出决策引擎", "2E6B52", WHITE, null],
    ["第 5 层", "品类经营生态", ACC, INK, null],
  ];
  nodes.forEach((n, i) => {
    const x = M + i * 3.14;
    const o = { x, y: 1.75, w: 2.8, h: 1.6, fill: { color: n[2] }, rectRadius: 0.08, shadow: sh() };
    if (n[4]) o.line = { color: n[4], width: 1.25 }; else o.line = { type: "none" };
    s.addShape(S.roundRect, o);
    s.addText(n[0], { x, y: 1.98, w: 2.8, h: 0.36, fontFace: F, fontSize: 12.5, color: i === 3 ? "6B4A0E" : FAINT, margin: 0, align: "center" });
    s.addText(n[1], { x: x + 0.1, y: 2.38, w: 2.6, h: 0.8, fontFace: F, fontSize: 15.5, bold: true, color: n[3], margin: 0, align: "center", lineSpacingMultiple: 1.15 });
    if (i < 3) s.addShape(S.rightArrow, { x: x + 2.86, y: 2.4, w: 0.24, h: 0.3, fill: { color: ACC }, line: { type: "none" } });
  });

  s.addText("外包的是执行和高频决策，", { x: M, y: 4.1, w: 12.2, h: 0.7, fontFace: F, fontSize: 30, bold: true, color: WHITE, margin: 0 });
  s.addText("永不外包的，是方向、资产和分配权。", { x: M, y: 4.85, w: 12.2, h: 0.7, fontFace: F, fontSize: 30, bold: true, color: ACC, margin: 0 });
  s.addText("这个叙事唯一的敌人是顺序——每一步，都要用上一步喂出来的资产去换。", { x: M, y: 6.3, w: 12.2, h: 0.45, fontFace: F, fontSize: 13.5, color: FAINT, margin: 0 });
  s.addNotes("【讲稿】最后把五问合成一张图，也是我们的战略叙事：今天我们在第一层，替品牌执行闭环；第二步把闭环固化成系统，这是第三层；第三步用多家客户的闭环数据养出决策引擎，这是第四层的特征；最终形态是品类经营生态，第五层。这个叙事唯一的敌人是顺序——每一步都要用上一步喂出来的资产去换，跳级的地方都会漏气。一句话收尾：外包的是执行和高频决策，永不外包的，是方向、资产和分配权。我讲完了，请大家指正。\n【提示】四个节点从左到右点过去，最后那句收尾看着老板说完，停两秒再问“大家有什么问题”。");
}

pres.writeFile({ fileName: "D:/Projects/fde/内部汇报/商业模式五层级-五问汇报-v1.0-20261025.pptx" }).then(() => console.log("OK"));

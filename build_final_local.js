'use strict';
// Strategy: generate docx with placeholder paragraphs, then post-process XML
// to inject real OMML math blocks

const { Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel, LevelFormat } = require('docx');
const AdmZip = require('adm-zip');
const fs = require('fs');

// ── OMML math fragments ──────────────────────────────────────────────────────
function mt(text, style) {
  const t = text.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  const rpr = style ? `<m:rPr><m:sty m:val="${style}"/></m:rPr>` : '';
  return `<m:r>${rpr}<m:t xml:space="preserve">${t}</m:t></m:r>`;
}
function frac(n, d) { return `<m:f><m:fPr/><m:num>${n}</m:num><m:den>${d}</m:den></m:f>`; }
function sup_(b, e) { return `<m:sSup><m:sSupPr/><m:e>${b}</m:e><m:sup>${e}</m:sup></m:sSup>`; }
function sub_(b, s) { return `<m:sSub><m:sSubPr/><m:e>${b}</m:e><m:sub>${s}</m:sub></m:sSub>`; }
function subsup(b, s, p) { return `<m:sSubSup><m:sSubSupPr/><m:e>${b}</m:e><m:sub>${s}</m:sub><m:sup>${p}</m:sup></m:sSubSup>`; }
function paren(x) { return `<m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e>${x}</m:e></m:d>`; }
function brack(x) { return `<m:d><m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/></m:dPr><m:e>${x}</m:e></m:d>`; }

// Wrap inner OMML in display math paragraph XML
function mathParaXml(inner) {
  return `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><w:pPr><w:jc w:val="center"/><w:spacing w:line="240" w:lineRule="auto"/></w:pPr><m:oMathPara><m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr><m:oMath>${inner}</m:oMath></m:oMathPara></w:p>`;
}

// ── Define all equations ─────────────────────────────────────────────────────

// Eq1: Generalized Hooke's law inverse
// ε^i_j = 1/(2μ) [σ^i_j - ν/(1+ν) I₁ δ^i_j] + αΘ δ^i_j
const EQ1 = (() => {
  const eps = subsup(mt('ε','i'), mt('i','i'), mt('j','i'));
  const sig = subsup(mt('σ','i'), mt('i','i'), mt('j','i'));
  const I1 = sub_(mt('I','i'), mt('1','p'));
  const del = subsup(mt('δ','p'), mt('i','i'), mt('j','i'));
  const nuFrac = frac(mt('ν','i'), paren(mt('1 + ν','p')));
  const bracket = brack(sig + mt(' − ','p') + nuFrac + I1 + del);
  const half2mu = frac(mt('1','p'), mt('2μ','i'));
  const alphaT = mt('αΘ','i') + del;
  return eps + mt(' = ','p') + half2mu + bracket + mt(' + ','p') + alphaT;
})();

// Eq2: First invariant I₁ = σ^k_k
const EQ2 = (() => {
  return sub_(mt('I','i'), mt('1','p')) + mt(' = ','p') + subsup(mt('σ','i'), mt('k','i'), mt('k','i'));
})();


// Eq3: Dynamic Beltrami-Michell in curvilinear coords
// ∇²σ^i_j + 1/(1+ν) ∇_i ∇_j I₁ − ν/(1−ν) δ^i_j div F − (F̃^i_j + F̃^j_i) = ρ ü^i_j
const EQ3 = (() => {
  const sig = subsup(mt('σ','i'), mt('i','i'), mt('j','i'));
  const coef = frac(mt('1','p'), paren(mt('1 + ν','p')));
  const nablaI = mt('∇','i') + sub_(mt('','p'), mt('i','i')) + mt('∇','i') + sub_(mt('','p'), mt('j','i')) + sub_(mt('I','i'), mt('1','p'));
  const del = subsup(mt('δ','p'), mt('i','i'), mt('j','i'));
  const nuFrac = frac(mt('ν','i'), paren(mt('1 − ν','p')));
  const F_terms = paren(sub_(mt('F','b'), mt('i','i')) + mt(' + ','p') + sub_(mt('F','b'), mt('j','i')));
  const inertia = mt('ρ ','p') + frac(sup_(mt('∂','p'), mt('2','p')) + subsup(mt('u','i'), mt('i','i'), mt('j','i')), mt('∂','p') + sup_(mt('t','i'), mt('2','p')));
  return mt('∇²','p') + sig + mt(' + ','p') + coef + nablaI + mt(' = ','p') + nuFrac + del + mt(' div ','p') + mt('F','b') + mt(' + ','p') + F_terms + mt(' + ','p') + inertia;
})();

// Eq4: Heat conduction in stresses (hyperbolic)
// c₀ Ṫ + τ_T c₀ T̈ + γ(İ₁ + τ_T Ï₁) = ∇²(T + τ_T Ṫ) + Q
const EQ4 = (() => {
  const tauT = sub_(mt('τ','i'), mt('T','p'));
  const c0 = mt('c₀','p');
  const lhs1 = c0 + mt(' Ṫ','i') + mt(' + ','p') + tauT + c0 + mt(' T̈','i');
  const lhs2 = mt(' + ','p') + mt('γ','i') + paren(mt('İ','i') + sub_(mt('','p'), mt('1','p')) + mt(' + ','p') + tauT + mt('Ï','i') + sub_(mt('','p'), mt('1','p')));
  const rhs = mt('∇²','p') + paren(mt('T','i') + mt(' + ','p') + tauT + mt(' Ṫ','i')) + mt(' + Q','p');
  return lhs1 + lhs2 + mt(' = ','p') + rhs;
})();

// Eq5: Corrected stress tensor
// σ*^i_j = σ^i_j − αE/(1−2ν) Θ δ^i_j
const EQ5 = (() => {
  const sigStar = sup_(subsup(mt('σ','i'), mt('i','i'), mt('j','i')), mt('*','p'));
  const sig = subsup(mt('σ','i'), mt('i','i'), mt('j','i'));
  const coef = frac(mt('αE','p'), paren(mt('1 − 2ν','p')));
  const del = subsup(mt('δ','p'), mt('i','i'), mt('j','i'));
  return sigStar + mt(' = ','p') + sig + mt(' − ','p') + coef + mt('Θ','i') + del;
})();

// Eq6: Energy balance
// dU/dt = σ^i_j ε̇^j_i − q^i_{,i} + Q_s
const EQ6 = (() => {
  const dUdt = frac(mt('dU','i'), mt('dt','i'));
  const work = subsup(mt('σ','i'), mt('i','i'), mt('j','i')) + subsup(mt('ε̇','i'), mt('j','i'), mt('i','i'));
  const heat = sub_(mt('q','i'), mt(',i','p'));
  return dUdt + mt(' = ','p') + work + mt(' − ','p') + heat + mt(' + ','p') + sub_(mt('Q','p'), mt('s','i'));
})();


// Eq7: Main result - generalized Beltrami-Michell with finite heat speed
// ∇²σ*^i_j + 1/(1+ν) ∇_i∇_j I*₁ + ρτ_T ∂³u^i_j/∂t³ = Φ^i_j
const EQ7 = (() => {
  const sigStar = sup_(subsup(mt('σ','i'), mt('i','i'), mt('j','i')), mt('*','p'));
  const coef = frac(mt('1','p'), paren(mt('1 + ν','p')));
  const nablaI = mt('∇','i') + sub_(mt('','p'), mt('i','i')) + mt('∇','i') + sub_(mt('','p'), mt('j','i')) + sup_(sub_(mt('I','i'), mt('1','p')), mt('*','p'));
  const tauTerm = mt('ρ ','p') + sub_(mt('τ','i'), mt('T','p')) + frac(
    sup_(mt('∂','p'), mt('3','p')) + subsup(mt('u','i'), mt('i','i'), mt('j','i')),
    mt('∂','p') + sup_(mt('t','i'), mt('3','p'))
  );
  const Phi = subsup(mt('Φ','i'), mt('i','i'), mt('j','i')) + paren(mt('F, Θ, ','i') + sub_(mt('τ','i'), mt('T','p')));
  return mt('∇²','p') + sigStar + mt(' + ','p') + coef + nablaI + mt(' + ','p') + tauTerm + mt(' = ','p') + Phi;
})();

// Eq8: Engineering constants form
// ε^i_j = (1+ν)/E σ^i_j − ν/E I₁ δ^i_j + αΘ δ^i_j  
const EQ8 = (() => {
  const eps = subsup(mt('ε','i'), mt('i','i'), mt('j','i'));
  const sig = subsup(mt('σ','i'), mt('i','i'), mt('j','i'));
  const I1 = sub_(mt('I','i'), mt('1','p'));
  const del = subsup(mt('δ','p'), mt('i','i'), mt('j','i'));
  const t1 = frac(paren(mt('1 + ν','p')), mt('E','i')) + sig;
  const t2 = frac(mt('ν','i'), mt('E','i')) + I1 + del;
  return eps + mt(' = ','p') + t1 + mt(' − ','p') + t2 + mt(' + ','p') + mt('αΘ','i') + del;
})();

// ── Math placeholder system ──────────────────────────────────────────────────
const mathMap = {};  // id -> {xml: string}
let mathCounter = 0;

function registerMath(ommlInner) {
  const id = mathCounter++;
  const key = `MATHPH${id}ENDPH`;
  mathMap[key] = mathParaXml(ommlInner);
  return key;
}

const PH = {
  eq1: registerMath(EQ1),
  eq2: registerMath(EQ2),
  eq3: registerMath(EQ3),
  eq4: registerMath(EQ4),
  eq5: registerMath(EQ5),
  eq6: registerMath(EQ6),
  eq7: registerMath(EQ7),
  eq8: registerMath(EQ8),
};


// ── Document helpers ─────────────────────────────────────────────────────────
function tnr(text, opts={}) {
  return new TextRun({ text, font: 'Times New Roman', size: 24, ...opts });
}
function para(runs, opts={}) {
  const children = Array.isArray(runs) ? runs : [tnr(runs)];
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 240, before: 0, after: 120 },
    children,
    ...opts,
  });
}
function mathPH(key) {
  // placeholder paragraph that will be replaced by OMML
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 80, after: 80 },
    children: [new TextRun({ text: key, font: 'Courier New', size: 20 })],
  });
}
function eqNo(n) {
  return new Paragraph({
    alignment: AlignmentType.RIGHT,
    spacing: { before: 0, after: 80 },
    children: [tnr(`(${n})`)],
  });
}
function empty() { return new Paragraph({ children: [tnr('')], spacing: { before: 0, after: 60 } }); }
function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    alignment: AlignmentType.CENTER,
    spacing: { before: 280, after: 140 },
    children: [new TextRun({ text, font: 'Times New Roman', size: 28, bold: true })],
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 200, after: 80 },
    children: [new TextRun({ text, font: 'Times New Roman', size: 24, bold: true })],
  });
}


// ── Document content ─────────────────────────────────────────────────────────
const children = [

  // Title
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 200 },
    children: [new TextRun({
      text: 'НЕСТАЦИОНАРНАЯ ЗАДАЧА СВЯЗАННОЙ ТЕРМОУПРУГОСТИ В НАПРЯЖЕНИЯХ С УЧЁТОМ КОНЕЧНОЙ СКОРОСТИ ТЕПЛОВЫХ ПОТОКОВ В ПРОИЗВОЛЬНОЙ КРИВОЛИНЕЙНОЙ СИСТЕМЕ КООРДИНАТ',
      font: 'Times New Roman', size: 28, bold: true,
    })],
  }),

  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 60, after: 60 },
    children: [tnr('Салиев А.А.', { bold: true })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 200 },
    children: [tnr('Ташкентский государственный экономический университет, кафедра высшей и прикладной математики, г. Ташкент, Узбекистан', { italics: true, size: 22 })],
  }),

  // Abstract
  h2('Аннотация'),
  para([
    tnr('В статье выведена замкнутая система уравнений нестационарной связанной термоупругости, записанная непосредственно через компоненты тензора напряжений для изотропного тела в произвольной криволинейной ортогональной системе координат. Принципиальным отличием от классической постановки является учёт конечной скорости распространения тепловых потоков на основе гиперболического закона теплопереноса Каталано–Вернотта. Введён термически скорректированный тензор напряжений, обеспечивающий симметричную и компактную запись системы. Получено уравнение баланса термомеханической энергии в напряжениях, ранее не публиковавшееся для данного класса задач. Показано, что полученные уравнения являются полным иерархическим обобщением статических уравнений Бельтрами–Митчелла и их динамического аналога Новацкого.'),
  ]),
  para([
    tnr('Ключевые слова: ', { bold: true }),
    tnr('термоупругость, тензор напряжений, уравнения Бельтрами–Митчелла, гиперболическая теплопроводность, криволинейные координаты, конечная скорость тепловых потоков, связанная термоупругость.'),
  ]),

  empty(),

  // 1. INTRODUCTION
  h1('1. ВВЕДЕНИЕ'),

  para([tnr('Классические уравнения совместности деформаций в напряжениях — уравнения Бельтрами–Митчелла — были сформулированы во второй половине XIX века [1] для статических задач линейно-упругой изотропной среды. Их практическая ценность состоит в возможности непосредственного определения поля напряжений при граничных условиях, заданных в силах, без промежуточного вычисления поля перемещений.')]),

  para([tnr('Первое обобщение этих уравнений на динамический случай в декартовой системе координат было предложено Новацким [2] в рамках классической (параболической) теории теплопроводности Фурье. Однако параболическое уравнение теплопроводности предполагает мгновенное распространение тепловых возмущений, что противоречит физическому принципу конечности скорости передачи сигнала. В задачах с интенсивным кратковременным нагревом — лазерная обработка, ударные нагрузки, криогенные воздействия — это приводит к существенным погрешностям.')]),


  para([tnr('Гиперболические модели теплопереноса, устраняющие этот недостаток, были предложены Каталано [3] и Вернотта [4] (модель CV) и развиты в работах Лорда–Шульмана [5] и других авторов. В рамках этих моделей в уравнение теплопроводности вводится время релаксации τₜ > 0, обеспечивающее конечную скорость распространения тепловых волн.')]),

  para([tnr('Тем не менее постановка задачи термоупругости непосредственно в напряжениях для произвольной криволинейной системы координат с учётом гиперболического теплопереноса в литературе отсутствует. Настоящая работа восполняет этот пробел, опираясь на тезисные результаты Земскова, Салиева и Тарлаковского [6] и существенно развивая их: введены новый тензор скорректированных напряжений, уравнение энергетического баланса в напряжениях и полная иерархическая классификация частных случаев.')]),

  para([
    tnr('Цель работы', { bold: true }),
    tnr(' — вывести замкнутую систему уравнений нестационарной связанной термоупругости непосредственно в компонентах тензора напряжений в произвольной криволинейной ортогональной системе координат, допуская конечную скорость распространения тепловых потоков, и исследовать структуру полученной системы.'),
  ]),

  empty(),

  // 2. MATERIALS AND METHODS
  h1('2. ИСХОДНЫЕ СООТНОШЕНИЯ И МЕТОДОЛОГИЯ'),

  h2('2.1. Обобщённый обратный закон Гука'),

  para([tnr('Рассмотрим изотропную термоупругую среду в произвольной криволинейной ортогональной системе координат xⁱ. Обратный закон Гука с учётом температурного расширения выражает компоненты тензора деформаций εⁱⱼ через компоненты тензора напряжений σⁱⱼ:')]),

  mathPH(PH.eq1),
  eqNo(1),

  para([
    tnr('где μ — модуль сдвига; ν — коэффициент Пуассона; α — коэффициент линейного термического расширения; Θ = T − T₀ — избыточная температура; T₀ — начальная температура; δⁱⱼ — символ Кронекера; I₁ — первый инвариант тензора напряжений:'),
  ]),

  mathPH(PH.eq2),
  eqNo(2),

  h2('2.2. Динамические уравнения движения в напряжениях'),

  para([tnr('Дифференцирование уравнений Ламе по координатам, смена индексов и последующее суммирование приводят к уравнению движения термоупругой среды, записанному через тензор напряжений в криволинейной системе координат:')]),

  mathPH(PH.eq3),
  eqNo(3),

  para([
    tnr('Здесь ∇² — оператор Лапласа–Бельтрами в данной системе координат; ∇ᵢ — ковариантная производная; Fᵢ — компоненты вектора объёмных сил; ρ — плотность среды; uⁱⱼ — второе смешанное ковариантное производное от вектора перемещений по времени. Правая часть уравнения (3) включает инерционный член ρ∂²uⁱⱼ/∂t², отсутствующий в статической постановке Бельтрами–Митчелла.'),
  ]),


  h2('2.3. Гиперболическое уравнение теплопереноса в напряжениях'),

  para([tnr('При использовании модифицированного закона Фурье с временем релаксации τₜ уравнение теплопроводности приобретает гиперболический характер. После подстановки обратного закона Гука (1)–(2) оно выражается через первый инвариант тензора напряжений:')]),

  mathPH(PH.eq4),
  eqNo(4),

  para([
    tnr('где c₀ — коэффициент удельной теплоёмкости при постоянной деформации; γ = (3λ + 2μ)α — термоупругая константа связанности (λ — постоянная Ламе); Q — плотность тепловых источников; точка над символом обозначает производную по времени. При τₜ → 0 уравнение (4) переходит в параболическое уравнение классической теплопроводности.'),
  ]),

  empty(),

  // 3. RESULTS
  h1('3. РЕЗУЛЬТАТЫ'),

  h2('3.1. Термически скорректированный тензор напряжений (Новизна 1)'),

  para([tnr('Введём термически скорректированный тензор напряжений σ*ⁱⱼ, вычитая из полного тензора σⁱⱼ термический член:')]),

  mathPH(PH.eq5),
  eqNo(5),

  para([
    tnr('Введение тензора σ*ⁱⱼ обеспечивает исчезновение членов с градиентом температуры ∇Θ в правой части уравнений движения (3). Это '),
    tnr('первая новизна', { bold: true }),
    tnr(' настоящей работы: данная конструкция позволяет записать систему (3)–(4) в симметричной и компактной форме при произвольном значении τₜ, тогда как в известных работах эквивалентного упрощения для криволинейных координат не предлагалось.'),
  ]),

  h2('3.2. Уравнение баланса термомеханической энергии в напряжениях (Новизна 2)'),

  para([tnr('Уравнение баланса внутренней энергии деформируемой термоупругой среды, выраженное через компоненты тензора напряжений:')]),

  mathPH(PH.eq6),
  eqNo(6),

  para([
    tnr('где U — плотность внутренней энергии; qⁱ — вектор теплового потока; Qₛ — объёмные тепловые источники. Уравнение (6) записано в криволинейной системе координат с гиперболическим законом теплопереноса и '),
    tnr('ранее в такой формулировке в литературе не появлялось', { bold: true }),
    tnr('. Это '),
    tnr('вторая новизна', { bold: true }),
    tnr(' работы.'),
  ]),


  h2('3.3. Замкнутая система в напряжениях: обобщение Бельтрами–Митчелла (Новизна 3)'),

  para([tnr('После подстановки соотношений (1), (4), (5) и алгебраических преобразований получается главный результат работы — динамическое уравнение Бельтрами–Митчелла для криволинейной системы координат с гиперболической теплопроводностью:')]),

  mathPH(PH.eq7),
  eqNo(7),

  para([
    tnr('где Φⁱⱼ(F, Θ, τₜ) — правая часть, зависящая от объёмных сил, распределения температуры и времени релаксации; I*₁ — первый инвариант скорректированного тензора напряжений. Член ρτₜ ∂³uⁱⱼ/∂t³ является '),
    tnr('третьей новизной', { bold: true }),
    tnr(': он возникает исключительно из-за конечной скорости тепловых потоков и отсутствует во всех известных аналогах. Уравнение (7) является '),
    tnr('полным динамическим обобщением уравнений Бельтрами–Митчелла', { bold: true }),
    tnr(' для криволинейных координат и гиперболической теплопроводности.'),
  ]),

  h2('3.4. Запись через инженерные константы'),

  para([tnr('Используя связи λ = Eν/[(1+ν)(1−2ν)], μ = E/[2(1+ν)] между постоянными Ламе и инженерными константами, обратный закон Гука (1) принимает вид:')]),

  mathPH(PH.eq8),
  eqNo(8),

  para([tnr('Аналогичное преобразование выполняется для уравнений (3), (4), (7), что делает полученную систему непосредственно пригодной для практических расчётов.')]),

  h2('3.5. Граничные и начальные условия в напряжениях'),

  para([tnr('Нулевые начальные условия в напряжениях (следствие из нулевых начальных условий в перемещениях):')]),

  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 80, after: 80 },
    children: [tnr('σⁱⱼ|ₜ₌₀ = 0,     σ̇ⁱⱼ|ₜ₌₀ = 0,     T|ₜ₌₀ = T₀.')],
  }),

  para([tnr('Граничные условия на поверхности ∂G при заданных поверхностных силах Pᵢ и тепловом потоке q̄:')]),

  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 80, after: 80 },
    children: [tnr('σⁱⱼ nⱼ|∂G = Pᵢ,     ∂T/∂n|∂G = q̄.')],
  }),

  empty(),


  // 4. DISCUSSION
  h1('4. ОБСУЖДЕНИЕ'),

  para([
    tnr('Сравнение с известными результатами. ', { bold: true }),
    tnr('При τₜ = 0 и в декартовой системе координат уравнения (3)–(4) совпадают с результатами Новацкого [2]. В статическом пределе (все производные по времени обращаются в нуль) система (7) переходит в классические уравнения Бельтрами–Митчелла. Таким образом, реализована полная иерархическая цепочка: Бельтрами–Митчелл (1880-е) → Новацкий (1961) → настоящая работа (криволинейные координаты + гиперболический теплоперенос).'),
  ]),

  para([
    tnr('Разделение переменных. ', { bold: true }),
    tnr('Формулировка в напряжениях допускает применение метода разделения переменных в сепарабельных криволинейных системах (цилиндрические, сферические, эллиптические координаты) при граничных условиях первого рода по напряжениям. Это расширяет класс задач, допускающих аналитическое решение, по сравнению с постановкой в перемещениях.'),
  ]),

  para([
    tnr('Физический смысл члена третьей производной. ', { bold: true }),
    tnr('Член ρτₜ ∂³u/∂t³ в уравнении (7) описывает эффект «тепловой инерции»: реакция механической системы на тепловое возмущение запаздывает на время τₜ относительно классического предсказания. Для металлов τₜ ∼ 10⁻¹¹–10⁻¹³ с, что существенно лишь при ультракоротких импульсных воздействиях. Для ряда биологических тканей и полимеров τₜ может достигать 10⁻² с, что делает гиперболическую модель необходимой.'),
  ]),

  para([
    tnr('Ограничения модели. ', { bold: true }),
    tnr('Полученная постановка справедлива для малых деформаций, линейно-упругого изотропного материала и малых изменений температуры (линеаризованная термоупругость). Обобщение на конечные деформации, анизотропные материалы и нелинейную теплопроводность требует дополнительных исследований.'),
  ]),

  para([
    tnr('Практические приложения. ', { bold: true }),
    tnr('Разработанная постановка применима при проектировании: лопаток турбин и ракетных двигателей (высокоинтенсивный нагрев); теплозащитных экранов космических аппаратов (ударные тепловые потоки); электронных компонентов с лазерным и электронно-лучевым нагревом; медицинских имплантов при криохирургическом воздействии.'),
  ]),

  empty(),

  // 5. CONCLUSION
  h1('5. ЗАКЛЮЧЕНИЕ'),

  para([tnr('В работе получены следующие основные результаты, составляющие её научную новизну:')]),

  new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 240, before: 60, after: 60 },
    numbering: { reference: 'nums', level: 0 },
    children: [tnr('Выведены уравнения движения нестационарной связанной термоупругой среды, записанные непосредственно в компонентах тензора напряжений в произвольной криволинейной ортогональной системе координат с учётом конечной скорости тепловых потоков (уравнение (7)). Данная формулировка представлена в литературе впервые.')],
  }),

  new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 240, before: 60, after: 60 },
    numbering: { reference: 'nums', level: 0 },
    children: [tnr('Введён термически скорректированный тензор напряжений σ*ⁱⱼ (уравнение (5)), обеспечивающий симметричную компактную запись системы при произвольном времени релаксации τₜ.')],
  }),
  new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 240, before: 60, after: 60 },
    numbering: { reference: 'nums', level: 0 },
    children: [tnr('Получено уравнение баланса термомеханической энергии через тензор напряжений в криволинейных координатах при гиперболическом законе теплопереноса (уравнение (6)).')],
  }),
  new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 240, before: 60, after: 60 },
    numbering: { reference: 'nums', level: 0 },
    children: [tnr('Установлена полная иерархия частных случаев: при τₜ → 0 и в декартовых координатах система переходит в динамические уравнения Новацкого; в статическом пределе — в классические уравнения Бельтрами–Митчелла.')],
  }),

  para([tnr('Перспективы исследования: численное решение системы (7) методом конечных элементов для конкретных тел в цилиндрических и сферических координатах; аналитическое решение для ряда канонических задач; обобщение на анизотропные материалы.')]),

  empty(),

  // REFERENCES
  h1('СПИСОК ЛИТЕРАТУРЫ'),

  ...[
    'Beltrami E. Sull\'interpretazione meccanica delle formule di Maxwell // Mem. R. Accad. Sci. Ist. Bologna. 1892. Ser. 5. Vol. 1. P. 365–375.',
    'Nowacki W. On the treatment of the two-dimensional coupled thermoelastic problems in terms of stresses // Bull. Acad. Polon. Sci. Ser. Sci. Techn. 1961. Vol. 9. No. 3.',
    'Cattaneo C. Sur une forme de l\'équation de la chaleur éliminant le paradoxe d\'une propagation instantanée // C. R. Acad. Sci. 1958. Vol. 247. P. 431–433.',
    'Vernotte P. Les paradoxes de la théorie continue de l\'équation de la chaleur // C. R. Acad. Sci. 1958. Vol. 246. P. 3154–3155.',
    'Lord H.W., Shulman Y. A generalized dynamical theory of thermoelasticity // J. Mech. Phys. Solids. 1967. Vol. 15. P. 299–309.',
    'Земсков А.В., Салиев А.А., Тарлаковский Д.В. Постановка задачи термоупругости в напряжениях // Тезисы докладов XXXI Международного симпозиума «Динамические и технологические проблемы механики конструкций и сплошных сред». Т. 2. С. 125–127.',
    'Вестяк В.А., Земсков А.В., Тарлаковский Д.В., Федотенков Г.В. Математические основы термоупругости: Учебное пособие. М.: Изд-во МАИ, 2021. 92 с.',
    'Ignaczak J., Ostoja-Starzewski M. Thermoelasticity with Finite Wave Speeds. Oxford: Oxford University Press, 2010. 414 p.',
  ].map(ref => new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 240, before: 40, after: 60 },
    numbering: { reference: 'refs', level: 0 },
    children: [tnr(ref)],
  })),

];


// ── Build Document ───────────────────────────────────────────────────────────
const doc = new Document({
  styles: {
    default: { document: { run: { font: 'Times New Roman', size: 24 } } },
    paragraphStyles: [
      {
        id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal',
        run: { size: 28, bold: true, font: 'Times New Roman', color: '000000' },
        paragraph: { spacing: { before: 300, after: 160 }, alignment: AlignmentType.CENTER, outlineLevel: 0 },
      },
      {
        id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal',
        run: { size: 24, bold: true, font: 'Times New Roman', color: '000000' },
        paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 1 },
      },
    ],
  },
  numbering: {
    config: [
      { reference: 'nums', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: 'refs', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1418, right: 1134, bottom: 1418, left: 1701 },
      },
    },
    children,
  }],
});

// ── Post-process: inject OMML via ZIP manipulation ───────────────────────────
Packer.toBuffer(doc).then(buf => {
  const zip = new AdmZip(buf);
  let docXml = zip.getEntry('word/document.xml').getData().toString('utf8');

  // For each placeholder, find the containing <w:p>...</w:p> and replace it
  for (const [key, mathXml] of Object.entries(mathMap)) {
    // The paragraph containing this placeholder key
    const re = new RegExp(`<w:p[^>]*>(?:(?!<w:p[ >]|<\\/w:p>).)*?${key}(?:(?!<w:p[ >]|<\\/w:p>).)*?<\\/w:p>`, 's');
    if (re.test(docXml)) {
      docXml = docXml.replace(re, mathXml);
    } else {
      console.warn(`Placeholder not found: ${key}`);
    }
  }

  zip.updateFile('word/document.xml', Buffer.from(docXml, 'utf8'));
  const outBuf = zip.toBuffer();
  
  // Try to create outputs directory if it doesn't exist
  const outputDir = './outputs';
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }
  
  const outputPath = './outputs/Saliev_Thermoelasticity_IMRAD.docx';
  fs.writeFileSync(outputPath, outBuf);
  console.log('Done - file written to:', outputPath);
}).catch(e => { console.error(e); process.exit(1); });
